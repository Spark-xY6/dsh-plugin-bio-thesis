import {spawn,spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const here=path.dirname(fileURLToPath(import.meta.url));
const release=path.resolve(process.env.THESIS_GIT_ROOT??path.join(here,'../../dsh-plugin-bio-thesis'));
const rev='4547e792ef7722b4056fdce57d0df4223b513084';
const skill='skills/bio-thesis-formatter/';
const hash=data=>createHash('sha256').update(data).digest('hex');
const freeze=JSON.parse(fs.readFileSync(path.join(here,'freeze.json')));
for(const [name,sha] of Object.entries(freeze)){
  if(hash(fs.readFileSync(path.join(here,name)))!==sha)throw Error('Freeze mismatch '+name);
}
const baseInput=path.join(here,'inputs');
const keys=Object.keys(freeze).filter(n=>n.startsWith('inputs/')).map(n=>n.slice(7)).sort();
if(keys.length!==5||fs.readdirSync(baseInput).sort().join('|')!==keys.join('|'))throw Error('Wrong input set');
const binary=process.env.CODEX_EXECUTABLE??path.join(process.env.APPDATA,'npm/node_modules/@openai/codex/node_modules/@openai/codex-win32-x64/vendor/x86_64-pc-windows-msvc/bin/codex.exe');
const python=process.env.THESIS_PYTHON??path.join(process.env.CONDA_PREFIX??'','python.exe');
if(!path.isAbsolute(python)||!fs.existsSync(python))throw Error('Set THESIS_PYTHON to the full path of the local Python interpreter');
const git=args=>{
  const p=spawnSync('git',['-C',release,...args],{maxBuffer:16*1024*1024,windowsHide:true});
  if(p.status!==0)throw Error('Cannot resolve official skill blobs');
  return p.stdout;
};
const runtime=git(['ls-tree','-r','--name-only',rev,'--',skill]).toString().trim().split(/\r?\n/);
const cli=['exec','--ignore-user-config','--ephemeral','--skip-git-repo-check','--enable','skip_host_skill_discovery','--disable','skill_search','--disable','plugins','--disable','apps','--disable','memories','--disable','multi_agent','-c','skills.bundled.enabled=false','-c','skills.include_instructions=false','-c','project_doc_max_bytes=0','-c','suppress_unstable_features_warning=true','-c','model_reasoning_effort="low"','-c','web_search="disabled"','-c','sandbox_workspace_write.network_access=false','-c','windows.sandbox="elevated"','-m','gpt-6.1-sol','-s','workspace-write','--json','-'];
const prompt=fs.readFileSync(path.join(here,'PROMPT.txt'),'utf8');
async function run(pair,arm){
  const logs=path.join(here,'trials','pair'+pair+'-'+arm);
  fs.mkdirSync(logs);
  const cwd=fs.mkdtempSync(path.join(os.tmpdir(),'campus-intake-'));
  fs.cpSync(baseInput,path.join(cwd,'inputs'),{recursive:true});
  fs.mkdirSync(path.join(cwd,'output'));
  if(arm==='skill')for(const p of runtime){
    if(!p.startsWith(skill)||p.includes('..'))throw Error('Invalid release path');
    const dest=path.join(cwd,'.agents',p);
    fs.mkdirSync(path.dirname(dest),{recursive:true});
    fs.writeFileSync(dest,git(['show',rev+':'+p]));
  }
  const env={...process.env,PYTHONDONTWRITEBYTECODE:'1',PYTHONUTF8:'1'};
  const pathkey=Object.keys(env).find(k=>k.toLowerCase()==='path')??'PATH';
  env[pathkey]=path.dirname(python)+';'+env[pathkey];
  for(const k of ['THESIS_GIT_ROOT','THESIS_PYTHON','CODEX_EXECUTABLE'])delete env[k];
  for(const k of ['CODEX_THREAD_ID','CODEX_SESSION_ID','CODEX_TOOL_CALL_ID','CODEX_APP_TOOLS_PIPE_PATH','CODEX_INTERNAL_ORIGINATOR_OVERRIDE','CODEX_WINDOWS_REGISTERED_CORE','CODEX_WINDOWS_SANDBOX_PACKAGE_FAMILY','CODEX_APP_RUNTIME_CLI_PATH'])delete env[k];
  const start=Date.now();
  fs.writeFileSync(path.join(logs,'started.json'),JSON.stringify({pair,arm,cwd,revision:rev,startUTC:new Date(start).toISOString(),cli,prompt,capSeconds:600},null,2));
  const stdout=path.join(logs,'events.jsonl');
  const fds=[fs.openSync(stdout,'wx'),fs.openSync(path.join(logs,'stderr.txt'),'wx')];
  const proc=spawn(binary,cli,{cwd,env,windowsHide:true,stdio:['pipe',...fds]});
  proc.stdin.on('error',()=>{});
  proc.stdin.end(prompt);
  let timedOut=false;
  const timer=setTimeout(()=>{timedOut=true;proc.kill()},600000);
  const exitCode=await new Promise((res,rej)=>{proc.on('error',rej);proc.on('close',res)});
  clearTimeout(timer);fds.forEach(n=>fs.closeSync(n));
  const secs=(Date.now()-start)/1000;
  const events=fs.readFileSync(stdout,'utf8').split(/\r?\n/).filter(Boolean).map(x=>JSON.parse(x));
  const items=events.filter(e=>e.type==='item.completed').map(e=>e.item);
  const toolCalls=items.filter(i=>!['reasoning','agent_message','error','todo_list'].includes(i.type)).length;
  const files=JSON.parse(fs.readFileSync(path.join(baseInput,'intake.json'))).map(r=>r.output);
  fs.mkdirSync(path.join(logs,'output'));
  for(const out of files)if(fs.existsSync(path.join(cwd,'output',out)))fs.copyFileSync(path.join(cwd,'output',out),path.join(logs,'output',out));
  const inputsIntact=fs.readdirSync(path.join(cwd,'inputs')).sort().join('|')===keys.join('|')&&keys.every(k=>hash(fs.readFileSync(path.join(cwd,'inputs',k)))===freeze['inputs/'+k]);
  const summary={pair,arm,elapsedSeconds:secs,timedOut,exitCode,toolCalls,usage:events.findLast(e=>e.type==='turn.completed')?.usage??null,outputFiles:fs.readdirSync(path.join(logs,'output')).sort(),inputsIntact};
  fs.writeFileSync(path.join(logs,'completed.json'),JSON.stringify(summary,null,2));
  console.log(JSON.stringify(summary));
}
fs.mkdirSync(path.join(here,'trials'));
await Promise.all([['bare','skill'],['skill','bare'],['bare','skill']].map(async (o,i)=>{for(const arm of o)await run(i+1,arm)}));
