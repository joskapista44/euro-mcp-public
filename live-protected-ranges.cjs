'use strict';

function status(status, operation, expected, actual, extra={}) {
  return {status, operation, source:'live-coedit-editor', humanObservationRequired:false, expected, actual, ...extra};
}

function commandBody(operation, spec={}) {
  const q=JSON.stringify;
  const sheet=`var s=${spec.sheetName?`Api.GetSheet(${q(spec.sheetName)})`:'Api.GetActiveSheet()'}; if(!s)return {ok:false,error:'sheet not found'};`;
  if(operation==='create') return `${sheet} var p=s.AddProtectedRange(${q(spec.title)},${q(spec.range)}); return {ok:!!p,count:(s.GetAllProtectedRanges()||[]).length,lookup:!!s.GetProtectedRange(${q(spec.title)})};`;
  if(operation==='inspect') return `${sheet} var p=s.GetProtectedRange(${q(spec.title)}); if(!p)return {ok:false,error:'protected range not found'}; function U(u){return u?{id:u.GetId(),name:u.GetName(),type:u.GetType()}:null} return {ok:true,count:(s.GetAllProtectedRanges()||[]).length,lookup:true,user:${spec.userId?`U(p.GetUser(${q(spec.userId)}))`:'null'},users:(p.GetAllUsers()||[]).map(U)};`;
  if(operation==='lookup-pair') return `${sheet} return {oldLookup:!!s.GetProtectedRange(${q(spec.title)}),newLookup:!!s.GetProtectedRange(${q(spec.newTitle)})};`;
  if(operation==='rename') return `${sheet} var p=s.GetProtectedRange(${q(spec.title)}); if(!p)return {ok:false,error:'protected range not found'}; p.SetTitle(${q(spec.newTitle)}); return true;`;
  if(operation==='add-user') return `${sheet} var p=s.GetProtectedRange(${q(spec.title)}); if(!p)return {ok:false,error:'protected range not found'}; function U(u){return u?{id:u.GetId(),name:u.GetName(),type:u.GetType()}:null} var a=p.AddUser(${q(spec.userId)},${q(spec.userName)},${q(spec.userType||'CanEdit')}); return {ok:!!a,added:U(a),getUser:U(p.GetUser(${q(spec.userId)})),users:(p.GetAllUsers()||[]).map(U)};`;
  if(operation==='delete-user') return `${sheet} var p=s.GetProtectedRange(${q(spec.title)}); if(!p)return {ok:false,error:'protected range not found'}; var r=p.DeleteUser(${q(spec.userId)}); return {ok:r===true,returnValue:r};`;
  throw new Error('unsupported protected range operation: '+operation);
}

async function call(frame, operation, spec={}) {
  const body=commandBody(operation,spec);
  return frame.evaluate(src=>new Promise(resolve=>window.Asc.editor.callCommand(new Function(src),false,resolve)),body);
}

async function execute(frame, operation, spec={}) {
  try {
    if(operation==='rename') {
      await call(frame,operation,spec);
      await new Promise(r=>setTimeout(r,spec.readbackDelayMs||1000));
      let readback=await call(frame,'lookup-pair',{sheetName:spec.sheetName,title:spec.title,newTitle:spec.newTitle});
      if(!readback){await new Promise(r=>setTimeout(r,500));readback=await call(frame,'lookup-pair',{sheetName:spec.sheetName,title:spec.title,newTitle:spec.newTitle});}
      const pass=readback&&readback.oldLookup===false&&readback.newLookup===true;
      return status(pass?'PASS':'FAIL',operation,{oldLookup:false,newLookup:true},{readback:readback||null});
    }
    const actual=await call(frame,operation,spec);
    if(operation==='create') return status(actual&&actual.ok&&actual.lookup?'PASS':'FAIL',operation,{lookup:true},actual);
    if(operation==='inspect') return status(actual&&actual.ok&&actual.lookup?'PASS':'FAIL',operation,{lookup:true},actual);
    if(operation==='add-user') {
      const u=actual&&actual.getUser;
      const pass=actual&&actual.ok&&u&&u.id===spec.userId&&u.name===spec.userName&&u.type===(spec.userType||'CanEdit');
      return status(pass?'PASS':'FAIL',operation,{id:spec.userId,name:spec.userName,type:spec.userType||'CanEdit'},actual);
    }
    if(operation==='delete-user') {
      if(!(actual&&actual.ok&&actual.returnValue===true)) return status('FAIL',operation,{deleteReturn:true},actual);
      await new Promise(r=>setTimeout(r,spec.readbackDelayMs||1500));
      const readback=await call(frame,'inspect',{sheetName:spec.sheetName,title:spec.title,userId:spec.userId});
      const pass=readback&&readback.ok&&readback.user===null&&!readback.users.some(u=>u&&u.id===spec.userId);
      return status(pass?'PASS':'FAIL',operation,{deleteReturn:true,user:null},{mutation:actual,readback});
    }
  } catch(e) { return status('FAIL',operation,'successful public API semantic readback',{error:String(e&&e.message?e.message:e)}); }
}

module.exports={execute,call,commandBody};
