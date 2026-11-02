import http from 'node:http';
import {readFile,stat,writeFile,mkdir,rename,chmod} from 'node:fs/promises';
import {extname,join,normalize,dirname} from 'node:path';
import {homedir} from 'node:os';
import {randomBytes,randomUUID,createHash,createPublicKey,verify} from 'node:crypto';
import ExcelJS from 'exceljs';

const port=Number(process.env.PORT||4173),root=process.cwd();
const chatgptModel=process.env.CHATGPT_MODEL||'gpt-5.6-luna';
const apiModel=process.env.AI_MODEL||'gpt-5.4-mini';
const apiBase=(process.env.AI_BASE_URL||'https://api.openai.com/v1').replace(/\/$/,'');
const configDir=join(process.env.LOCALAPPDATA||homedir(), 'DashboardStudio');
const authPath=join(configDir,'chatgpt-auth.json');
const deepseekPath=join(configDir,'deepseek-config.json');
const databasePath=join(configDir,'database-config.json');
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json'};
const send=(res,status,body,type='application/json',headers={})=>{res.writeHead(status,{'Content-Type':type,'Cache-Control':'no-store',...headers});res.end(type==='application/json'?JSON.stringify(body):body)};
const readBody=req=>new Promise((resolve,reject)=>{let body='',settled=false;req.on('data',c=>{if(settled)return;body+=c;if(body.length>15_000_000){settled=true;reject(new Error('请求过大，请使用 8MB 以内的文件'))}});req.on('end',()=>{if(settled)return;try{resolve(JSON.parse(body||'{}'))}catch{reject(new Error('JSON 格式错误'))}});req.on('error',reject)});
const b64url=buf=>Buffer.from(buf).toString('base64url');
const pendingAuth=new Map();
let cachedAuth;
let cachedDeepSeek;
let cachedDatabase;
let lastAuthError=null;

function remoteError(payload,fallback='远程服务请求失败'){
  if(!payload)return fallback;if(typeof payload==='string')return payload;
  const error=payload.error,detail=payload.detail;
  if(typeof payload.error_description==='string')return payload.error_description;
  if(typeof error==='string')return error;
  if(error&&typeof error==='object')return [error.code||error.type,error.message||error.detail].filter(Boolean).join('：')||fallback;
  if(typeof detail==='string')return detail;
  if(detail&&typeof detail==='object')return [detail.code||detail.type,detail.message||detail.detail].filter(Boolean).join('：')||fallback;
  if(typeof payload.message==='string')return payload.message;return fallback;
}

const metricSchema={type:'object',additionalProperties:false,required:['dashboard'],properties:{dashboard:{type:'object',additionalProperties:false,required:['title','rationale','metrics'],properties:{title:{type:'string'},rationale:{type:'string'},metrics:{type:'array',minItems:4,maxItems:12,items:{type:'object',additionalProperties:false,required:['id','name','desc','chart','icon','status'],properties:{id:{type:'string'},name:{type:'string'},desc:{type:'string'},chart:{type:'string',enum:['指标卡','指标卡 + 趋势','折线图','柱状图','饼图','漏斗图','数据表']},icon:{type:'string'},status:{type:'string',enum:['draft']}}}}}}}};
const dataParameterSchema={type:'object',additionalProperties:false,required:['key','label','role','dataType','required','description'],properties:{key:{type:'string'},label:{type:'string'},role:{type:'string',enum:['measure','dimension','date','filter','join']},dataType:{type:'string'},required:{type:'boolean'},description:{type:'string'}}};
const dataRequirementItemSchema={type:'object',additionalProperties:false,required:['metricId','metricName','grain','aggregation','parameters','sqlTemplate'],properties:{metricId:{type:'string'},metricName:{type:'string'},grain:{type:'string'},aggregation:{type:'string'},parameters:{type:'array',minItems:1,maxItems:10,items:dataParameterSchema},sqlTemplate:{type:'string'}}};
const visualDashboardSchema={type:'object',additionalProperties:false,required:['dashboard'],properties:{dashboard:{type:'object',additionalProperties:false,required:['title','subtitle','rationale','theme','components'],properties:{title:{type:'string'},subtitle:{type:'string'},rationale:{type:'string'},theme:{type:'object',additionalProperties:false,required:['background','surface','text','muted','accent','secondary'],properties:{background:{type:'string'},surface:{type:'string'},text:{type:'string'},muted:{type:'string'},accent:{type:'string'},secondary:{type:'string'}}},components:{type:'array',minItems:1,maxItems:24,items:{type:'object',additionalProperties:false,required:['id','metricId','title','type','span','displayValue','note','labels','values'],properties:{id:{type:'string'},metricId:{type:'string'},title:{type:'string'},type:{type:'string',enum:['kpi','line','bar','donut','table']},span:{type:'integer',enum:[3,4,6,8,12]},displayValue:{type:'string'},note:{type:'string'},labels:{type:'array',maxItems:6,items:{type:'string'}},values:{type:'array',maxItems:6,items:{type:'number'}}}}}}}}};
const dataRequirementsSchema={type:'object',additionalProperties:false,required:['dataRequirements'],properties:{dataRequirements:{type:'array',minItems:1,maxItems:24,items:dataRequirementItemSchema}}};
const sqlCompletionSchema={type:'object',additionalProperties:false,required:['sql','explanation','assumptions','usedTables','usedFields'],properties:{sql:{type:'string'},explanation:{type:'string'},assumptions:{type:'array',maxItems:8,items:{type:'string'}},usedTables:{type:'array',maxItems:12,items:{type:'string'}},usedFields:{type:'array',maxItems:40,items:{type:'string'}}}};
const datasetSchema={type:'object',additionalProperties:false,required:['dataset'],properties:{dataset:{type:'object',additionalProperties:false,required:['summary','tables','relationships'],properties:{summary:{type:'string'},tables:{type:'array',items:{type:'object',additionalProperties:false,required:['name','businessName','description','rowCount','primaryKey','fields'],properties:{name:{type:'string'},businessName:{type:'string'},description:{type:'string'},rowCount:{type:'integer'},primaryKey:{type:'string'},fields:{type:'array',items:{type:'object',additionalProperties:false,required:['name','type','meaning','sample'],properties:{name:{type:'string'},type:{type:'string'},meaning:{type:'string'},sample:{type:'string'}}}}}}},relationships:{type:'array',items:{type:'object',additionalProperties:false,required:['from','to','reason'],properties:{from:{type:'string'},to:{type:'string'},reason:{type:'string'}}}}}}}};

async function loadAuth(){if(cachedAuth!==undefined)return cachedAuth;try{cachedAuth=JSON.parse(await readFile(authPath,'utf8'))}catch{cachedAuth=null}return cachedAuth}
async function saveAuth(value){await mkdir(dirname(authPath),{recursive:true});const temp=`${authPath}.tmp`;await writeFile(temp,JSON.stringify(value,null,2),{mode:0o600});try{await chmod(temp,0o600)}catch{}await rename(temp,authPath);cachedAuth=value}
async function loadDeepSeek(){if(cachedDeepSeek!==undefined)return cachedDeepSeek;try{cachedDeepSeek=JSON.parse(await readFile(deepseekPath,'utf8'))}catch{cachedDeepSeek=null}return cachedDeepSeek}
async function saveDeepSeek(value){await mkdir(dirname(deepseekPath),{recursive:true});const temp=`${deepseekPath}.tmp`;await writeFile(temp,JSON.stringify(value,null,2),{mode:0o600});try{await chmod(temp,0o600)}catch{}await rename(temp,deepseekPath);cachedDeepSeek=value}
async function deepseekSettings(){const saved=await loadDeepSeek();return {apiKey:process.env.DEEPSEEK_API_KEY||saved?.apiKey||'',model:process.env.DEEPSEEK_MODEL||saved?.model||'deepseek-flash'}}
async function loadDatabase(){if(cachedDatabase!==undefined)return cachedDatabase;try{cachedDatabase=JSON.parse(await readFile(databasePath,'utf8'))}catch{cachedDatabase={mode:'offline'}}return cachedDatabase}
async function saveDatabase(value){await mkdir(dirname(databasePath),{recursive:true});const temp=`${databasePath}.tmp`;await writeFile(temp,JSON.stringify(value,null,2),{mode:0o600});try{await chmod(temp,0o600)}catch{}await rename(temp,databasePath);cachedDatabase=value}
function publicDatabaseConfig(config){return {mode:config?.mode||'offline',configured:Boolean(config?.mode==='online'&&config?.host&&config?.user&&config?.lastTestedAt),provider:config?.provider||'mysql',host:config?.host||'',port:config?.port||null,database:config?.database||'',schema:config?.schema||'',user:config?.user||'',encrypt:Boolean(config?.encrypt),trustServerCertificate:config?.trustServerCertificate!==false,stats:config?.lastSyncedAt?config?.stats||null:null,lastTestedAt:config?.lastTestedAt||null,lastSyncedAt:config?.lastSyncedAt||null,passwordSaved:Boolean(config?.password)}}
function decodePart(value){return JSON.parse(Buffer.from(value,'base64url').toString('utf8'))}
async function validateIdToken(token,clientId,nonce){
  const [encodedHeader,encodedPayload,encodedSignature]=token.split('.');if(!encodedSignature)throw new Error('ID token 格式无效');
  const header=decodePart(encodedHeader),payload=decodePart(encodedPayload);if(header.alg!=='RS256')throw new Error(`不支持的 ID token 算法：${header.alg}`);
  const discovery=await fetch('https://auth.openai.com/.well-known/openid-configuration').then(r=>r.json());
  const keys=await fetch(discovery.jwks_uri).then(r=>r.json());const jwk=keys.keys?.find(k=>k.kid===header.kid);if(!jwk)throw new Error('找不到对应的 OpenAI 签名密钥');
  const valid=verify('RSA-SHA256',Buffer.from(`${encodedHeader}.${encodedPayload}`),createPublicKey({key:jwk,format:'jwk'}),Buffer.from(encodedSignature,'base64url'));if(!valid)throw new Error('ID token 签名验证失败');
  const audiences=Array.isArray(payload.aud)?payload.aud:[payload.aud];if(payload.iss!==discovery.issuer||!audiences.includes(clientId)||payload.exp*1000<=Date.now()||payload.nonce!==nonce)throw new Error('ID token 身份校验失败');return payload;
}
async function beginChatGPTAuth(res){
  const existing=await loadAuth(),state=b64url(randomBytes(24)),nonce=b64url(randomBytes(24)),verifier=b64url(randomBytes(48));
  const challenge=b64url(createHash('sha256').update(verifier).digest()),hostId=existing?.ext_agent_host_id||`urn:uuid:${randomUUID()}`,clientId=existing?.client_id||'dynamic_agent_client';
  const redirectUri=`http://127.0.0.1:${port}/auth/callback`,params=new URLSearchParams({client_id:clientId,response_type:'code',redirect_uri:redirectUri,scope:'openid profile email offline_access resource.invoke chatgpt.tokens.use.direct',resource:'https://api.openai.com/v1',state,nonce,code_challenge_method:'S256',code_challenge:challenge,ext_agent_host_id:hostId});
  if(clientId==='dynamic_agent_client')params.set('agent_name_hint','Dashboard Studio');else if(existing?.id_token)params.set('id_token_hint',existing.id_token);
  if(!existing?.scopes?.includes('chatgpt.tokens.use.direct'))params.set('prompt','consent');
  lastAuthError=null;
  pendingAuth.set(state,{nonce,verifier,redirectUri,clientId,hostId,createdAt:Date.now()});send(res,302,'', 'text/plain', {Location:`https://auth.openai.com/api/accounts/authorize?${params}`});
}
async function completeChatGPTAuth(url,res){
  const state=url.searchParams.get('state'),pending=pendingAuth.get(state);if(!pending||Date.now()-pending.createdAt>10*60*1000)throw new Error('登录请求已过期，请重新开始');pendingAuth.delete(state);
  if(url.searchParams.get('error'))throw new Error(`授权未完成：${url.searchParams.get('error_description')||url.searchParams.get('error')}`);
  const code=url.searchParams.get('code'),issuedClient=url.searchParams.get('client_id')||pending.clientId;if(!code||issuedClient==='dynamic_agent_client')throw new Error('没有收到有效的授权代码或客户端 ID');
  if(pending.clientId!=='dynamic_agent_client'&&issuedClient!==pending.clientId)throw new Error('返回的客户端 ID 不匹配');
  const prior=await loadAuth();await saveAuth({...prior,client_id:issuedClient,ext_agent_host_id:pending.hostId,scopes:prior?.scopes||[]});
  const form=new URLSearchParams({grant_type:'authorization_code',client_id:issuedClient,code,code_verifier:pending.verifier,redirect_uri:pending.redirectUri,resource:'https://api.openai.com/v1'});
  const tokenRes=await fetch('https://auth.openai.com/api/accounts/oauth/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:form});const tokens=await tokenRes.json();if(!tokenRes.ok)throw new Error(`令牌交换失败（HTTP ${tokenRes.status}）：${remoteError(tokens)}`);
  const identity=await validateIdToken(tokens.id_token,issuedClient,pending.nonce),scopes=String(tokens.scope||url.searchParams.get('scope')||'').split(/\s+/).filter(Boolean);if(!scopes.includes('chatgpt.tokens.use.direct'))throw new Error('账户没有授权 ChatGPT 会员推理额度');
  await saveAuth({email:identity.email||null,subject:identity.sub,issuer:identity.iss,client_id:issuedClient,ext_agent_host_id:pending.hostId,id_token:tokens.id_token,access_token:tokens.access_token,refresh_token:tokens.refresh_token,token_type:tokens.token_type||'Bearer',expires_in:tokens.expires_in||3600,scopes,saved_at:new Date().toISOString()});lastAuthError=null;
  send(res,302,'','text/plain',{Location:'/?auth=success'});
}
async function getChatGPTToken(){
  const auth=await loadAuth();if(!auth?.access_token||!auth.scopes?.includes('chatgpt.tokens.use.direct'))throw new Error('请先使用 ChatGPT 登录');
  const expiresAt=new Date(auth.saved_at).getTime()+Number(auth.expires_in||3600)*1000;if(Date.now()<expiresAt-60_000)return auth.access_token;if(!auth.refresh_token)throw new Error('ChatGPT 登录已过期，请重新授权');
  const form=new URLSearchParams({grant_type:'refresh_token',client_id:auth.client_id,refresh_token:auth.refresh_token,resource:'https://api.openai.com/v1'});const response=await fetch('https://auth.openai.com/api/accounts/oauth/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:form});const next=await response.json();if(!response.ok)throw new Error('ChatGPT 会话刷新失败，请重新授权');
  await saveAuth({...auth,access_token:next.access_token,refresh_token:next.refresh_token||auth.refresh_token,id_token:next.id_token||auth.id_token,expires_in:next.expires_in||3600,scopes:String(next.scope||auth.scopes.join(' ')).split(/\s+/),saved_at:new Date().toISOString()});return next.access_token;
}

function extractOutput(data){return data.output_text||data.output?.flatMap(x=>x.content||[]).find(x=>x.type==='output_text')?.text}
function parseStructuredOutput(raw,label='模型'){
  const text=String(raw||'').trim();
  if(!text)throw new Error(`${label}没有返回结构化结果`);
  const unfenced=text.replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'').trim();
  const candidates=[unfenced];
  const first=unfenced.indexOf('{'),last=unfenced.lastIndexOf('}');
  if(first>=0&&last>first)candidates.push(unfenced.slice(first,last+1));
  let parseError;
  for(const candidate of [...new Set(candidates)]){try{return JSON.parse(candidate)}catch(error){parseError=error}}
  const message=String(parseError?.message||'');
  if(/unterminated|unexpected end|end of json/i.test(message)||last<first)throw new Error(`${label}返回内容被截断`);
  throw new Error(`${label}返回的 JSON 格式不完整`);
}
const databaseDefaults={mysql:3306,sqlserver:1433,oracle:1521,dm:5236};
function normalizeDatabaseConfig(input,existing={}){
  const mode=input.mode==='online'?'online':'offline',provider=['mysql','sqlserver','oracle','dm'].includes(input.provider)?input.provider:(existing.provider||'mysql');
  if(mode==='offline')return {...existing,mode};
  const config={...existing,mode,provider,host:String(input.host||existing.host||'').trim(),port:Number(input.port||existing.port||databaseDefaults[provider]),database:String(input.database||existing.database||'').trim(),schema:String(input.schema||existing.schema||'').trim(),user:String(input.user||existing.user||'').trim(),password:String(input.password||existing.password||''),encrypt:Boolean(input.encrypt),trustServerCertificate:input.trustServerCertificate!==false};
  if(!config.host||!config.user||!config.password)throw new Error('请填写主机、用户名和密码');
  if(provider!=='dm'&&!config.database)throw new Error('请填写数据库名或 Oracle Service Name');
  if(!Number.isInteger(config.port)||config.port<1||config.port>65535)throw new Error('数据库端口无效');
  if(config.schema&&!/^[\w$#]+$/u.test(config.schema))throw new Error('Schema 只能包含字母、数字、下划线、$ 或 #');
  return config;
}
async function executeDatabaseQuery(config,sqlText,maxRows=500){
  if(config.provider==='mysql'){
    const mysql=await import('mysql2/promise'),connection=await mysql.createConnection({host:config.host,port:config.port,user:config.user,password:config.password,database:config.database||undefined,ssl:config.encrypt?{rejectUnauthorized:!config.trustServerCertificate}:undefined,connectTimeout:8000});
    try{const [rows,fields]=await connection.query({sql:sqlText,rowsAsArray:false});return {rows:Array.isArray(rows)?rows.slice(0,maxRows):[],columns:(fields||[]).map(field=>field.name)}}finally{await connection.end()}
  }
  if(config.provider==='sqlserver'){
    const imported=await import('mssql'),mssql=imported.default||imported,pool=await new mssql.ConnectionPool({server:config.host,port:config.port,user:config.user,password:config.password,database:config.database||undefined,connectionTimeout:8000,requestTimeout:12000,options:{encrypt:config.encrypt,trustServerCertificate:config.trustServerCertificate}}).connect();
    try{const result=await pool.request().query(sqlText),rows=(result.recordset||[]).slice(0,maxRows);return {rows,columns:Object.keys(result.recordset?.columns||rows[0]||{})}}finally{await pool.close()}
  }
  if(config.provider==='oracle'){
    const imported=await import('oracledb'),oracledb=imported.default||imported,connection=await oracledb.getConnection({user:config.user,password:config.password,connectString:`${config.host}:${config.port}/${config.database}`});
    try{const result=await connection.execute(sqlText,[],{outFormat:oracledb.OUT_FORMAT_OBJECT,maxRows});return {rows:result.rows||[],columns:(result.metaData||[]).map(column=>column.name)}}finally{await connection.close()}
  }
  const imported=await import('dmdb'),dmdb=imported.default||imported,connection=await dmdb.getConnection({connectString:`${config.host}:${config.port}`,user:config.user,password:config.password,schema:config.schema||config.user,socketTimeout:12000});
  try{const result=await connection.execute(sqlText,[],{outFormat:dmdb.OUT_FORMAT_OBJECT,maxRows});return {rows:result.rows||[],columns:(result.metaData||[]).map(column=>column.name)}}finally{await connection.close()}
}
function firstScalar(result){const row=result.rows?.[0]||{},value=Object.values(row)[0];return Number(value??0)}
function qualifiedSystemTable(config,name){return config.schema?`${config.schema}.${name}`:name}
async function databaseStats(config){
  const tableResult=await executeDatabaseQuery(config,`SELECT COUNT(*) AS item_count FROM ${qualifiedSystemTable(config,'ctl_tbl')}`,1);
  const fieldResult=await executeDatabaseQuery(config,`SELECT COUNT(*) AS item_count FROM ${qualifiedSystemTable(config,'fg_columns')}`,1);
  return {tableCount:firstScalar(tableResult),fieldCount:firstScalar(fieldResult),source:'database'};
}
async function testDatabaseConnection(config){
  const sqlText=config.provider==='oracle'?'SELECT 1 AS connection_ok FROM dual':'SELECT 1 AS connection_ok';
  await executeDatabaseQuery(config,sqlText,1);
}
function quoteColumn(provider,name){if(provider==='mysql')return `\`${String(name).replace(/`/g,'``')}\``;if(provider==='sqlserver')return `[${String(name).replace(/]/g,']]')}]`;return `"${String(name).replace(/"/g,'""')}"`}
function applyRowLimit(provider,selectSql,limit){if(provider==='sqlserver')return selectSql.replace(/^SELECT\s+/i,`SELECT TOP ${limit} `);if(provider==='mysql')return `${selectSql} LIMIT ${limit}`;return `SELECT * FROM (${selectSql}) scoped_rows WHERE ROWNUM <= ${limit}`}
function castText(provider,column){const quoted=quoteColumn(provider,column);if(provider==='mysql')return `LOWER(CAST(${quoted} AS CHAR))`;if(provider==='sqlserver')return `LOWER(CAST(${quoted} AS NVARCHAR(4000)))`;return `LOWER(CAST(${quoted} AS VARCHAR2(4000)))`}
function metricSearchTerms(metric){const text=`${metric?.name||''} ${metric?.desc||''}`.replace(/\s+/g,' '),terms=new Set(),name=String(metric?.name||'').trim();if(name)terms.add(name.slice(0,24));const base=name.replace(/(同比|环比|趋势|分布|明细|分析|指标|金额|数量|占比|比率|率|数)$/,'');if(base.length>=2)terms.add(base);for(const part of text.split(/[\s,，。；;、/()（）+→-]+/)){const clean=part.trim();if(clean.length>=2&&clean.length<=12)terms.add(clean)}return [...terms].slice(0,6)}
function scoreColumn(name,kind){const value=String(name).toLowerCase();if(kind==='table'){if(/^(table_name|tablename|tblname|c_tablename|c_name)$/.test(value))return 100;if(/table.*name|tbl.*name|belong.*table/.test(value))return 80;if(value==='name')return 40}if(kind==='tableRef'){if(/^(table_name|tablename|tblname|c_tablename|c_belongtable)$/.test(value))return 100;if(/table|tbl|belong/.test(value))return 70}if(kind==='field'){if(/^(field_name|column_name|colname|c_name)$/.test(value))return 100;if(/field.*name|column.*name|col.*name/.test(value))return 80;if(value==='name')return 40}if(kind==='type'&&/type|datatype|data_type|c_type/.test(value))return 70;if(kind==='meaning'&&/fullname|caption|title|meaning|remark|description|desc|memo|c_fullname/.test(value))return 70;return 0}
function pickColumn(columns,kind){return [...columns].sort((a,b)=>scoreColumn(b,kind)-scoreColumn(a,kind)).find(column=>scoreColumn(column,kind)>0)||null}
function searchableColumns(columns){const matched=columns.filter(column=>/name|table|tbl|field|column|caption|title|fullname|remark|desc|memo|code/i.test(column));return (matched.length?matched:columns).slice(0,8)}
function likeWhere(provider,columns,terms){const safeTerms=terms.map(term=>String(term).toLowerCase().replace(/'/g,"''"));return columns.flatMap(column=>safeTerms.map(term=>`${castText(provider,column)} LIKE '%${term}%'`)).join(' OR ')}
async function inspectSystemTable(config,tableName){const table=qualifiedSystemTable(config,tableName),result=await executeDatabaseQuery(config,applyRowLimit(config.provider,`SELECT * FROM ${table}`,1),1);return {table,columns:result.columns,rows:result.rows}}
async function synchronizeDatabaseSchema(config){
  const stats=await databaseStats(config),ctl=await inspectSystemTable(config,'ctl_tbl'),columnsInfo=await inspectSystemTable(config,'fg_columns');
  const tableNameColumn=pickColumn(ctl.columns,'table'),businessColumn=pickColumn(ctl.columns,'meaning'),tableRefColumn=pickColumn(columnsInfo.columns,'tableRef'),fieldNameColumn=pickColumn(columnsInfo.columns,'field'),typeColumn=pickColumn(columnsInfo.columns,'type'),meaningColumn=pickColumn(columnsInfo.columns,'meaning');
  if(!tableNameColumn)throw new Error('无法识别 ctl_tbl 中的表名字段');
  if(!tableRefColumn||!fieldNameColumn)throw new Error('无法识别 fg_columns 中的所属表或字段名列');
  const ctlRows=(await executeDatabaseQuery(config,applyRowLimit(config.provider,`SELECT * FROM ${ctl.table}`,200),200)).rows;
  const catalog=ctlRows.map(row=>({name:String(row[tableNameColumn]??'').trim(),businessName:businessColumn?String(row[businessColumn]??'').trim():'',description:'',rowCount:0,primaryKey:'',fields:[]})).filter(table=>/^[\w$#]+$/u.test(table.name));
  const previewNames=catalog.slice(0,80).map(table=>table.name),fieldRows=[];
  for(let offset=0;offset<previewNames.length;offset+=20){
    const names=previewNames.slice(offset,offset+20).map(name=>`'${name.toLowerCase().replace(/'/g,"''")}'`).join(',');
    const where=`${castText(config.provider,tableRefColumn)} IN (${names})`;
    fieldRows.push(...(await executeDatabaseQuery(config,applyRowLimit(config.provider,`SELECT * FROM ${columnsInfo.table} WHERE ${where}`,800),800)).rows);
  }
  const byName=new Map(catalog.slice(0,80).map(table=>[table.name.toLowerCase(),table]));
  for(const row of fieldRows){
    const table=byName.get(String(row[tableRefColumn]??'').trim().toLowerCase()),name=String(row[fieldNameColumn]??'').trim();
    if(!table||!name||table.fields.length>=24)continue;
    table.fields.push({name,type:typeColumn?String(row[typeColumn]??''):'',meaning:meaningColumn?String(row[meaningColumn]??''):'',sample:''});
  }
  return {summary:`已从数据库同步 ${stats.tableCount} 张表、${stats.fieldCount} 个字段；页面保留前 ${Math.min(catalog.length,80)} 张表的轻量摘要，AI SQL 补全仍按指标实时检索。`,tables:catalog.slice(0,80),relationships:[],stats,truncated:stats.tableCount>catalog.length||stats.fieldCount>fieldRows.length};
}
async function targetedDatabaseSchema(config,metric){
  const terms=metricSearchTerms(metric),ctl=await inspectSystemTable(config,'ctl_tbl'),ctlSearch=searchableColumns(ctl.columns),ctlWhere=likeWhere(config.provider,ctlSearch,terms),ctlRows=(await executeDatabaseQuery(config,applyRowLimit(config.provider,`SELECT * FROM ${ctl.table}${ctlWhere?` WHERE ${ctlWhere}`:''}`,16),16)).rows,tableNameColumn=pickColumn(ctl.columns,'table');
  const candidateNames=[...new Set(ctlRows.map(row=>String(row[tableNameColumn]??'').trim()).filter(name=>/^[\w$#]+$/u.test(name)))].slice(0,12),businessColumn=pickColumn(ctl.columns,'meaning');
  const columnsInfo=await inspectSystemTable(config,'fg_columns'),tableRefColumn=pickColumn(columnsInfo.columns,'tableRef'),fieldNameColumn=pickColumn(columnsInfo.columns,'field'),typeColumn=pickColumn(columnsInfo.columns,'type'),meaningColumn=pickColumn(columnsInfo.columns,'meaning');
  let fieldWhere='';if(candidateNames.length&&tableRefColumn){const names=candidateNames.map(name=>`'${name.toLowerCase().replace(/'/g,"''")}'`).join(',');fieldWhere=`${castText(config.provider,tableRefColumn)} IN (${names})`}else fieldWhere=likeWhere(config.provider,searchableColumns(columnsInfo.columns),terms);
  const fieldRows=(await executeDatabaseQuery(config,applyRowLimit(config.provider,`SELECT * FROM ${columnsInfo.table}${fieldWhere?` WHERE ${fieldWhere}`:''}`,300),300)).rows,tables=new Map();
  for(const name of candidateNames){const source=ctlRows.find(row=>String(row[tableNameColumn]??'').trim()===name);tables.set(name,{name,businessName:businessColumn?String(source?.[businessColumn]??''):'',description:'',fields:[]})}
  for(const row of fieldRows){const tableName=String(row[tableRefColumn]??candidateNames[0]??'').trim();if(!tableName)continue;if(!tables.has(tableName))tables.set(tableName,{name:tableName,businessName:'',description:'',fields:[]});const fieldName=String(row[fieldNameColumn]??'').trim();if(fieldName)tables.get(tableName).fields.push({name:fieldName,type:typeColumn?String(row[typeColumn]??''):'',meaning:meaningColumn?String(row[meaningColumn]??''):''})}
  return {summary:`按关键词 ${terms.join('、')} 检索 ctl_tbl，命中 ${tables.size} 张候选表；仅读取这些表在 fg_columns 中的字段。`,tables:[...tables.values()].filter(table=>table.fields.length).slice(0,12),relationships:[],searchTerms:terms,scannedAll:false};
}
function targetedOfflineSchema(schema,metric){const terms=metricSearchTerms(metric),tables=(schema?.tables||[]).map(table=>{const haystack=JSON.stringify(table).toLowerCase(),score=terms.reduce((sum,term)=>sum+(haystack.includes(term.toLowerCase())?1:0),0);return {table,score}}).filter(item=>item.score>0).sort((a,b)=>b.score-a.score).slice(0,12).map(item=>({...item.table,fields:(item.table.fields||[]).slice(0,100)}));return {summary:`离线元数据按关键词 ${terms.join('、')} 筛选，共 ${tables.length} 张候选表。`,tables,relationships:(schema?.relationships||[]).filter(relation=>tables.some(table=>String(relation.from||'').startsWith(table.name))||tables.some(table=>String(relation.to||'').startsWith(table.name))),searchTerms:terms,scannedAll:false}}
function cellText(cell){const value=cell?.value;if(value==null)return '';if(value instanceof Date)return value.toISOString();if(typeof value==='object'){if('result' in value)return String(value.result??'');if(Array.isArray(value.richText))return value.richText.map(x=>x.text||'').join('');if('text' in value)return String(value.text??'')}return String(value)}
function sheetSnapshot(worksheet){const maxColumn=Math.min(Math.max(worksheet.columnCount,1),80),searchRows=Math.min(Math.max(worksheet.rowCount,1),10);let headerRow=1,bestCount=-1;for(let rowNumber=1;rowNumber<=searchRows;rowNumber++){let count=0;for(let column=1;column<=maxColumn;column++)if(cellText(worksheet.getCell(rowNumber,column)).trim())count++;if(count>bestCount){bestCount=count;headerRow=rowNumber}}const headers=[];for(let column=1;column<=maxColumn;column++){const name=cellText(worksheet.getCell(headerRow,column)).trim();if(name||column<=bestCount)headers.push({column,name:name||`column_${column}`})}const samples=[];for(let rowNumber=headerRow+1;rowNumber<=Math.min(worksheet.rowCount,headerRow+5);rowNumber++){const row={};for(const header of headers)row[header.name]=cellText(worksheet.getCell(rowNumber,header.column)).slice(0,160);if(Object.values(row).some(Boolean))samples.push(row)}return {sheetName:worksheet.name,rowCount:Math.max(0,worksheet.rowCount-headerRow),headers:headers.map(x=>x.name),samples}}
function parseCsvLine(line){const values=[];let value='',quoted=false;for(let i=0;i<line.length;i++){const char=line[i];if(char==='"'){if(quoted&&line[i+1]==='"'){value+='"';i++}else quoted=!quoted}else if(char===','&&!quoted){values.push(value);value=''}else value+=char}values.push(value);return values}
async function extractWorkbook(fileName,buffer){if(/\.csv$/i.test(fileName)){const lines=buffer.toString('utf8').replace(/^\uFEFF/,'').split(/\r?\n/).filter(Boolean),headers=parseCsvLine(lines[0]||'').map((x,i)=>x.trim()||`column_${i+1}`),samples=lines.slice(1,6).map(line=>Object.fromEntries(headers.map((header,i)=>[header,parseCsvLine(line)[i]?.slice(0,160)||''])));return [{sheetName:fileName.replace(/\.csv$/i,''),rowCount:Math.max(0,lines.length-1),headers:headers.slice(0,80),samples}]}const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(buffer);return workbook.worksheets.slice(0,20).map(sheetSnapshot)}
async function analyzeDataSource(input){const settings=await deepseekSettings();if(!settings.apiKey)throw new Error('请先配置 DeepSeek API Key，再识别 Excel');const fileName=String(input.fileName||'').trim();if(!/\.(xlsx|csv)$/i.test(fileName))throw new Error('仅支持 .xlsx 或 .csv 文件');const buffer=Buffer.from(String(input.dataBase64||''),'base64');if(!buffer.length)throw new Error('文件内容为空');if(buffer.length>8*1024*1024)throw new Error('文件超过 8MB');const workbook=await extractWorkbook(fileName,buffer);if(!workbook.length)throw new Error('文件中没有可识别的工作表');const instructions='你是企业数据建模专家。根据 Excel 工作表名、表头和少量样例，识别每张业务表、字段数据类型、字段业务含义、可能的主键与跨表关联。必须保留原始工作表名和原始字段名；不得虚构文件中不存在的字段；样例仅用于推断，输出中不要泄露超过给定样例的信息。';const format={type:'json_schema',name:'dataset_analysis',strict:true,schema:datasetSchema},response=await fetch('https://api.deepseek.com/responses',{method:'POST',headers:{Authorization:`Bearer ${settings.apiKey}`,'Content-Type':'application/json'},body:JSON.stringify({model:input.model||settings.model,instructions,input:JSON.stringify({businessContext:String(input.businessContext||'').slice(0,4000),workbook}),text:{format},max_output_tokens:6000})}),data=await response.json();if(!response.ok)throw new Error(remoteError(data,'DeepSeek 数据识别失败'));const text=extractOutput(data);if(!text)throw new Error('DeepSeek 没有返回数据结构');return JSON.parse(text)}
async function streamOutput(response){
  if(!response.ok){const err=await response.text();throw new Error(err||'ChatGPT 请求失败')}
  const raw=await response.text();let output='',completed=false;for(const line of raw.split(/\r?\n/)){if(!line.startsWith('data:'))continue;const value=line.slice(5).trim();if(!value||value==='[DONE]')continue;try{const event=JSON.parse(value);if(event.type==='response.output_text.delta')output+=event.delta||'';if(event.type==='response.completed')completed=true;if(event.type==='response.failed')throw new Error(event.response?.error?.message||event.response?.error?.code||'生成失败')}catch(error){if(error instanceof SyntaxError)continue;throw error}}
  if(!completed)throw new Error('ChatGPT 响应未完整结束');return output;
}
async function generateDashboard(input){
  const instructions=`你是企业数据看板架构师。输出简洁的中文看板原型方案，指标口径清楚且适合管理层。${input.prototypeMode?'当前处于概念原型阶段：允许根据业务场景、提示词和参考图提出指标，但必须把尚未绑定数据库字段的内容标记为草稿，不得声称已经取到真实数据。':'当前已提供数据库摘要：指标必须遵循给出的真实表和字段，不得编造字段。'}重新生成时必须在布局或指标组合上给出明显不同的版本。`;
  const format={type:'json_schema',name:'dashboard_plan',strict:true,schema:metricSchema};
  const references=Array.isArray(input.referenceImages)?input.referenceImages.slice(0,3):[],textInput={...input,referenceImages:references.map(r=>({name:r.name,attached:true}))};
  const content=[{type:'input_text',text:JSON.stringify(textInput)},...references.filter(r=>typeof r.dataUrl==='string'&&r.dataUrl.startsWith('data:image/')).map(r=>({type:'input_image',image_url:r.dataUrl,detail:'low'}))];
  const modelInput=[{role:'user',content}];
  if(input.authMode==='deepseek'){
    const settings=await deepseekSettings();if(!settings.apiKey)throw new Error('请先在 AI 设置中保存 DeepSeek API Key');
    const response=await fetch('https://api.deepseek.com/responses',{method:'POST',headers:{Authorization:`Bearer ${settings.apiKey}`,'Content-Type':'application/json'},body:JSON.stringify({model:input.model||settings.model,instructions,input:modelInput,text:{format},max_output_tokens:3000})});
    const data=await response.json();if(!response.ok)throw new Error(remoteError(data,'DeepSeek 请求失败'));const text=extractOutput(data);if(!text)throw new Error('DeepSeek 没有返回结构化结果');return JSON.parse(text);
  }
  if(input.authMode==='chatgpt'){
    const token=await getChatGPTToken();const response=await fetch('https://api.openai.com/v1/responses',{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({model:input.model||chatgptModel,instructions,input:modelInput,store:false,stream:true,text:{format}})});const text=await streamOutput(response);return JSON.parse(text);
  }
  const key=process.env.AI_API_KEY||process.env.OPENAI_API_KEY;if(!key)throw new Error('未配置 AI_API_KEY');const response=await fetch(`${apiBase}/responses`,{method:'POST',headers:{Authorization:`Bearer ${key}`,'Content-Type':'application/json'},body:JSON.stringify({model:input.model||apiModel,instructions,input:modelInput,text:{format}})});const data=await response.json();if(!response.ok)throw new Error(data.error?.message||'API 模型请求失败');const text=extractOutput(data);if(!text)throw new Error('模型没有返回结构化结果');return JSON.parse(text);
}
async function requestDeepSeekJson({settings,model,instructions,content,format,label,maxOutputTokens}){
  let lastError;
  for(let attempt=0;attempt<2;attempt++){
    const retryNote=attempt?' 上一次输出未完整闭合；请进一步压缩文字，只返回完整合法的 JSON。':'';
    const response=await fetch('https://api.deepseek.com/responses',{method:'POST',headers:{Authorization:`Bearer ${settings.apiKey}`,'Content-Type':'application/json'},body:JSON.stringify({model:model||settings.model,instructions:`${instructions}${retryNote}`,input:[{role:'user',content}],text:{format},max_output_tokens:maxOutputTokens})});
    const data=await response.json();if(!response.ok)throw new Error(remoteError(data,`${label}生成失败`));
    try{return parseStructuredOutput(extractOutput(data),label)}catch(error){lastError=error;if(!/截断|JSON/.test(error.message))throw error}
  }
  throw new Error(`${lastError?.message||`${label}生成失败`}，已自动重试，请稍后再试`);
}
async function completeMetricSql(input){
  const settings=await deepseekSettings();if(!settings.apiKey)throw new Error('请先配置 DeepSeek API Key');
  const database=await loadDatabase(),metric=input.metric||{},online=database.mode==='online'&&database.host,context=online?await targetedDatabaseSchema(database,metric):targetedOfflineSchema(input.databaseSummary||null,metric);
  if(!context.tables?.length)throw new Error(online?'没有在 ctl_tbl / fg_columns 中检索到相关候选表，请补充更明确的指标说明':'离线元数据中没有匹配的候选表，请先上传相关表字段');
  const dialect=input.dialect||({mysql:'MySQL',sqlserver:'SQL Server',oracle:'Oracle',dm:'DM8'}[database.provider]||'SQL Server');
  const instructions='你是熟悉企业系统数据库的 SQL 助手。只允许使用输入 candidateSchema 中存在的表和字段生成指标取数 SQL。ctl_tbl 与 fg_columns 的信息可能不完全准确；p_form 或 p_form_ 开头通常是自定义表单，主表常以 _m 结尾，明细表常以 _d/_d1 结尾，常见主明细关联为主表.phid=明细表.pphid，但必须以已提供字段为准。表头与表体存在同名业务字段时不得自行猜测，应在 assumptions 中说明歧义，并在 SQL 注释中标记待确认。帮助字段不得擅自转换中文；只有上下文明确提供帮助关联时才可关联。多个一对多来源必须先分别按目标粒度预汇总再关联，避免金额放大。不得 SELECT *，不得引用候选上下文以外的字段。输出与 dialect 一致的可执行 SQL；若信息不足，仍输出带 TODO 注释的安全草稿并列出假设。';
  const payload={metric,currentSql:String(input.currentSql||'').slice(0,12000),dialect,requirement:input.requirement||null,candidateSchema:context};
  const result=await requestDeepSeekJson({settings,model:input.model,instructions,content:[{type:'input_text',text:JSON.stringify(payload)}],format:{type:'json_schema',name:'sql_completion',strict:true,schema:sqlCompletionSchema},label:'DeepSeek SQL 补全',maxOutputTokens:3500});
  return {...result,schemaContext:context,source:online?'database':'offline',scope:{candidateTableCount:context.tables.length,candidateFieldCount:context.tables.reduce((sum,table)=>sum+(table.fields?.length||0),0),searchTerms:context.searchTerms}};
}
async function generateVisualDashboard(input){
  if(!Array.isArray(input.metrics)||!input.metrics.length)throw new Error('请先完成指标大纲，再生成看板');
  const settings=await deepseekSettings();if(!settings.apiKey)throw new Error('请先在 AI 设置中保存 DeepSeek API Key');
  const references=Array.isArray(input.referenceImages)?input.referenceImages.slice(0,3):[],textInput={metrics:input.metrics,databaseSummary:input.databaseSummary||null,businessContext:input.businessContext||'',prompt:input.prompt||'',styleGuide:input.styleGuide||'',project:input.project||'',referenceImages:references.map(r=>({name:r.name,attached:true}))};
  const layoutContent=[{type:'input_text',text:JSON.stringify(textInput)},...references.filter(r=>typeof r.dataUrl==='string'&&r.dataUrl.startsWith('data:image/')).map(r=>({type:'input_image',image_url:r.dataUrl,detail:'low'}))];
  const requirementContent=[{type:'input_text',text:JSON.stringify({metrics:input.metrics,databaseSummary:input.databaseSummary||null,businessContext:input.businessContext||''})}];
  const layoutInstructions='你是资深企业数据看板设计师。只输出可视看板布局，不输出 SQL 或参数。metrics 是最高优先级：每个指标必须在 components 中出现，metricId 原样使用，名称口径不得改变。databaseSummary 用于判断图表维度；businessContext、prompt、styleGuide 和参考图决定配色、明暗、信息密度与布局。没有真实值时 displayValue 必须为“--”；values 仅作样式预览，note 必须写“样式预览，待接入真实数据”。颜色使用六位十六进制，12 列栅格，每个组件最多 6 个标签和值，文字简洁。';
  const requirementInstructions='你是数据分析工程师。只输出每个 metrics 指标的取数参数契约，不输出看板布局。每个 metricId 必须原样出现一次。结合 databaseSummary 说明统计粒度、聚合方式和最少必要参数，参数说明不超过 30 个汉字。sqlTemplate 使用 {{参数key}} 表示待映射字段；FROM 主表用 {{source_table}}；从参数字段推导表名用 {{table:参数key}}，例如 SELECT SUM({{amount_field}}) FROM {{table:amount_field}}。不得编造固定字段或固定表名，SQL 保持简短。';
  const [layout,requirements]=await Promise.all([
    requestDeepSeekJson({settings,model:input.model,instructions:layoutInstructions,content:layoutContent,format:{type:'json_schema',name:'visual_dashboard',strict:true,schema:visualDashboardSchema},label:'DeepSeek 看板布局',maxOutputTokens:4500}),
    requestDeepSeekJson({settings,model:input.model,instructions:requirementInstructions,content:requirementContent,format:{type:'json_schema',name:'data_requirements',strict:true,schema:dataRequirementsSchema},label:'DeepSeek 取数参数',maxOutputTokens:6000})
  ]);
  return {dashboard:{...layout.dashboard,dataRequirements:requirements.dataRequirements}};
}

const server=http.createServer(async(req,res)=>{
  const url=new URL(req.url,`http://127.0.0.1:${port}`);
  if(url.pathname==='/api/status'){
    const auth=await loadAuth(),deepseek=await deepseekSettings(),database=await loadDatabase();return send(res,200,{deepseek:{configured:Boolean(deepseek.apiKey),label:'DeepSeek API',model:deepseek.model,keyHint:deepseek.apiKey?`尾号 ${deepseek.apiKey.slice(-4)}`:null},database:publicDatabaseConfig(database),chatgpt:{configured:Boolean(auth?.access_token&&auth.scopes?.includes('chatgpt.tokens.use.direct')),identityConnected:Boolean(auth?.client_id),label:'ChatGPT Plus',model:chatgptModel,email:auth?.email||null},api:{configured:Boolean(process.env.AI_API_KEY||process.env.OPENAI_API_KEY),label:'兼容 API',model:apiModel,baseUrl:apiBase},lastAuthError});
  }
  if(url.pathname==='/api/database/status'&&req.method==='GET')return send(res,200,publicDatabaseConfig(await loadDatabase()));
  if(url.pathname==='/api/database/settings'&&req.method==='POST'){try{const body=await readBody(req),existing=await loadDatabase(),config=normalizeDatabaseConfig(body,existing);if(config.mode==='offline'){await saveDatabase(config);return send(res,200,publicDatabaseConfig(config))}await testDatabaseConnection(config);const saved={...config,stats:null,lastTestedAt:new Date().toISOString(),lastSyncedAt:null};await saveDatabase(saved);return send(res,200,publicDatabaseConfig(saved))}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/database/sync-schema'&&req.method==='POST'){try{const config=await loadDatabase();if(config.mode!=='online'||!config.lastTestedAt)throw new Error('请先在数据库连接页完成连接测试');const dataset=await synchronizeDatabaseSchema(config),saved={...config,stats:dataset.stats,lastSyncedAt:new Date().toISOString()};await saveDatabase(saved);return send(res,200,{dataset,database:publicDatabaseConfig(saved)})}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/settings/deepseek'&&req.method==='POST'){try{const body=await readBody(req),existing=await deepseekSettings(),apiKey=String(body.apiKey||existing.apiKey||'').trim(),model=String(body.model||existing.model||'deepseek-flash').trim();if(!apiKey)throw new Error('请输入 DeepSeek API Key');if(apiKey.length<16)throw new Error('API Key 格式似乎不完整');await saveDeepSeek({apiKey,model,savedAt:new Date().toISOString()});return send(res,200,{configured:true,model,keyHint:`尾号 ${apiKey.slice(-4)}`})}catch(error){return send(res,400,{error:error.message})}}
  if(url.pathname==='/api/test-ai'&&req.method==='POST'){try{const body=await readBody(req);if(body.provider!=='deepseek')throw new Error('当前仅支持测试 DeepSeek 连接');const settings=await deepseekSettings();if(!settings.apiKey)throw new Error('尚未配置 DeepSeek API Key');const response=await fetch('https://api.deepseek.com/responses',{method:'POST',headers:{Authorization:`Bearer ${settings.apiKey}`,'Content-Type':'application/json'},body:JSON.stringify({model:settings.model,input:'只回复 OK',max_output_tokens:8})});const data=await response.json();if(!response.ok)throw new Error(remoteError(data,'DeepSeek 连接失败'));return send(res,200,{ok:true,model:settings.model})}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/analyze-data-source'&&req.method==='POST'){try{const result=await analyzeDataSource(await readBody(req)),fieldCount=(result.dataset?.tables||[]).reduce((sum,table)=>sum+(table.fields?.length||0),0);return send(res,200,{...result,fieldCount})}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/auth/chatgpt/start'&&req.method==='GET')return beginChatGPTAuth(res);
  if(url.pathname==='/auth/callback'&&req.method==='GET'){try{return await completeChatGPTAuth(url,res)}catch(error){lastAuthError=error.message;return send(res,400,`<!doctype html><meta charset="utf-8"><title>授权失败</title><style>body{font:16px sans-serif;padding:40px;color:#24332d}a{color:#076b50}</style><h2>ChatGPT 授权未完成</h2><p>${String(error.message).replace(/[<>&]/g,'')}</p><p>返回设置后可重新授权，API Key 方式仍可继续使用。</p><a href="/">返回看板工场</a>`,'text/html; charset=utf-8')}}
  if(url.pathname==='/api/generate-dashboard'&&req.method==='POST'){try{return send(res,200,await generateDashboard(await readBody(req)))}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/generate-visual-dashboard'&&req.method==='POST'){try{return send(res,200,await generateVisualDashboard(await readBody(req)))}catch(error){return send(res,503,{error:error.message})}}
  if(url.pathname==='/api/complete-sql'&&req.method==='POST'){try{return send(res,200,await completeMetricSql(await readBody(req)))}catch(error){return send(res,503,{error:error.message})}}
  if(req.method!=='GET')return send(res,405,{error:'Method not allowed'});
  const raw=url.pathname==='/'?'/index.html':decodeURIComponent(url.pathname),path=normalize(join(root,raw));if(!path.startsWith(root))return send(res,403,{error:'Forbidden'});
  try{const info=await stat(path);if(!info.isFile())throw new Error();send(res,200,await readFile(path),types[extname(path)]||'application/octet-stream')}catch{send(res,404,{error:'Not found'})}
});
server.listen(port,'127.0.0.1',async()=>{const deepseek=await deepseekSettings();console.log(`看板工场已启动：http://127.0.0.1:${port}\nDeepSeek API：${deepseek.apiKey?'已配置':'未配置'}\n备用 API：${process.env.AI_API_KEY||process.env.OPENAI_API_KEY?'已配置':'未配置'}`)});
