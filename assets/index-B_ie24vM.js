(function(){const o=document.createElement("link").relList;if(o&&o.supports&&o.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))a(n);new MutationObserver(n=>{for(const r of n)if(r.type==="childList")for(const d of r.addedNodes)d.tagName==="LINK"&&d.rel==="modulepreload"&&a(d)}).observe(document,{childList:!0,subtree:!0});function i(n){const r={};return n.integrity&&(r.integrity=n.integrity),n.referrerPolicy&&(r.referrerPolicy=n.referrerPolicy),n.crossOrigin==="use-credentials"?r.credentials="include":n.crossOrigin==="anonymous"?r.credentials="omit":r.credentials="same-origin",r}function a(n){if(n.ep)return;n.ep=!0;const r=i(n);fetch(n.href,r)}})();function ve(s){let o=null,i=null;return s.SheetNames.forEach(a=>{a.includes("台北")?o=a:a.includes("新北")&&(i=a)}),{taipeiSheetName:o,newTaipeiSheetName:i}}function w(s){if(!s)return-1;let o=0;const i=s.toUpperCase();for(let a=0;a<i.length;a++)o=o*26+i.charCodeAt(a)-64;return o-1}function D(s,o,i,a=null){const n={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!s||s.length===0)return n;const r=w(o.category||""),d=w(o.gender||""),m=w(o.age||""),y=w(o.living||""),t=w(o.cms||"");for(let c=1;c<s.length;c++){const p=s[c];if((p[r]||"").toString().trim()!==i)continue;const v=parseInt(p[m]);if(!(a&&(isNaN(v)||a.min!==void 0&&v<a.min||a.max!==void 0&&v>a.max))){if(d!==-1&&p[d]!==void 0){const l=p[d].toString().trim();l==="男"?n.gender.男++:l==="女"&&n.gender.女++}if(isNaN(v)||(v<=49?n.age["<= 49"]++:v>=50&&v<=64?n.age["50 >= && <= 64"]++:v>=65&&v<=74?n.age["65 >= && <= 74"]++:v>=75&&v<=84?n.age["75 >= && <= 84"]++:v>=85&&n.age["85 >="]++),y!==-1&&p[y]!==void 0){let l=p[y].toString().trim();l&&(l==="其他類型居住狀況"&&(l="其他"),n.living[l]=(n.living[l]||0)+1)}if(t!==-1&&p[t]!==void 0){const l=p[t].toString().trim();l&&(n.cms[l]=(n.cms[l]||0)+1)}}}return n}function ue(s,o,i,a){const n=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],r=(t,c)=>{const p=[[`【${c}】`,""],["男",t.gender.男||0],["女",t.gender.女||0],["",""]];return p.push(["【年齡分佈】",""],["<= 49",t.age["<= 49"]||0],["50 >= && <= 64",t.age["50 >= && <= 64"]||0],["65 >= && <= 74",t.age["65 >= && <= 74"]||0],["75 >= && <= 84",t.age["75 >= && <= 84"]||0],["85 >=",t.age["85 >="]||0],["",""]),p},d=t=>{const c=["1","2","3","4","5","6","7","8"],p=Object.keys(t.cms).filter(l=>!c.includes(l)).sort(),P=c.concat(p),v=[["",""],["【CMS 等級】",""]];return P.forEach(l=>{t.cms[l]!==void 0&&v.push([l,t.cms[l]])}),v},m=()=>{let t=r(s,"性別統計");t.push(["(65歲以上 老人)",""]),n.forEach(v=>{t.push([v==="與家人或其他人同住"?"與家人住":v,o.living[v]||0])}),t.push(["",""],["(50~64 身障)",""]),["獨居","僅配偶2人","與家人住","與朋友住","其他"].forEach(v=>{let l=0;v==="僅配偶2人"?l=i.living["獨居(兩老)"]||0:v==="與家人住"?l=i.living.與家人或其他人同住||0:l=i.living[v]||0,t.push([v,l])}),t.push(["",""],["身心障礙者",""]);const p=Object.values(i.gender).reduce((v,l)=>v+l,0),P=Object.values(o.gender).reduce((v,l)=>v+l,0);return t.push(["50-64歲",p],["65歲以上",P]),t=t.concat(d(s)),t},y=()=>{let t=r(a,"性別統計");return t.push(["【居住狀況】",""]),n.forEach(c=>{a.living[c]!==void 0&&t.push([c,a.living[c]])}),t=t.concat(d(a)),t};return{elderlyRows:m(),disabledRows:y()}}function pe(s,o,i,a,n,r){const d=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],m=Object.values(i.gender).reduce((c,p)=>c+p,0),y=Object.values(o.gender).reduce((c,p)=>c+p,0),t=document.createElement("div");return t.className="preview-group",t.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">台北老福</span>
            <span class="preview-group-badge">${n}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${s.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${s.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${s.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${s.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${s.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${s.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${s.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(s.cms).sort().map(c=>`<div class="stat-item"><span>${c}</span> <span class="stat-value">${s.cms[c]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
        <div class="stats-grid triples mt-4">
            <div class="stat-card">
                <div class="stat-label">(65歲以上 老人)</div>
                <div class="stat-values">
                    ${d.map(c=>`<div class="stat-item"><span>${c==="與家人或其他人同住"?"與家人住":c}</span> <span class="stat-value">${o.living[c]||0}</span></div>`).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">(50~64 身障)</div>
                <div class="stat-values">
                    ${["獨居","僅配偶2人","與家人住","與朋友住","其他"].map(c=>{let p=0;return c==="僅配偶2人"?p=i.living["獨居(兩老)"]||0:c==="與家人住"?p=i.living.與家人或其他人同住||0:p=i.living[c]||0,`<div class="stat-item"><span>${c}</span> <span class="stat-value">${p}</span></div>`}).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">身心障礙者</div>
                <div class="stat-values">
                    <div class="stat-item"><span>50-64歲</span> <span class="stat-value">${m}</span></div>
                    <div class="stat-item"><span>65歲以上</span> <span class="stat-value">${y}</span></div>
                </div>
            </div>
        </div>

        <div class="preview-group-header mt-8">
            <span class="preview-group-title">台北身障</span>
            <span class="preview-group-badge">${r}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${a.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${a.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${a.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${a.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${a.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${a.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${a.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">居住狀況</div>
                <div class="stat-values">
                    ${d.filter(c=>a.living[c]!==void 0).map(c=>`<div class="stat-item"><span>${c}</span> <span class="stat-value">${a.living[c]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(a.cms).sort().map(c=>`<div class="stat-item"><span>${c}</span> <span class="stat-value">${a.cms[c]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
    `,t}function me(s,o){const i={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!s||s.length===0)return i;const a=w(o.gender||""),n=w(o.age||""),r=w(o.living||""),d=w(o.cms||"");for(let m=1;m<s.length;m++){const y=s[m];if(a!==-1&&y[a]!==void 0){const t=y[a].toString().trim();t==="男"?i.gender.男++:t==="女"&&i.gender.女++}if(n!==-1&&y[n]!==void 0){const t=parseInt(y[n]);isNaN(t)||(t<=49?i.age["<= 49"]++:t>=50&&t<=64?i.age["50 >= && <= 64"]++:t>=65&&t<=74?i.age["65 >= && <= 74"]++:t>=75&&t<=84?i.age["75 >= && <= 84"]++:t>=85&&i.age["85 >="]++)}if(r!==-1&&y[r]!==void 0){let t=y[r].toString().trim();t&&(t==="其他類型居住狀況"&&(t="其他"),i.living[t]=(i.living[t]||0)+1)}if(d!==-1&&y[d]!==void 0){const t=y[d].toString().trim();t&&(i.cms[t]=(i.cms[t]||0)+1)}}return i}function ge(s){const o=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],i=["1","2","3","4","5","6","7","8"],a=Object.keys(s.cms).filter(d=>!i.includes(d)).sort(),n=i.concat(a),r=[["【性別統計】",""],["男",s.gender.男||0],["女",s.gender.女||0],["",""]];return r.push(["【年齡分佈】",""],["<= 49",s.age["<= 49"]||0],["50 >= && <= 64",s.age["50 >= && <= 64"]||0],["65 >= && <= 74",s.age["65 >= && <= 74"]||0],["75 >= && <= 84",s.age["75 >= && <= 84"]||0],["85 >=",s.age["85 >="]||0],["",""]),r.push(["【居住狀況】",""]),o.forEach(d=>{s.living[d]!==void 0&&r.push([d,s.living[d]])}),r.push(["",""],["【CMS 等級】",""]),n.forEach(d=>{s.cms[d]!==void 0&&r.push([d,s.cms[d]])}),r}function fe(s,o){const i=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],a=["1","2","3","4","5","6","7","8"],n=Object.keys(s.cms).filter(m=>!a.includes(m)).sort(),r=a.concat(n),d=document.createElement("div");return d.className="preview-group",d.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">新北</span>
            <span class="preview-group-badge">${o}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${s.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${s.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${s.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${s.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${s.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${s.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${s.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">居住狀況</div>
                <div class="stat-values">
                    ${i.filter(m=>s.living[m]!==void 0).map(m=>`<div class="stat-item"><span>${m}</span> <span class="stat-value">${s.living[m]}</span></div>`).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${r.filter(m=>s.cms[m]!==void 0).map(m=>`<div class="stat-item"><span>${m}</span> <span class="stat-value">${s.cms[m]}</span></div>`).join("")}
                </div>
            </div>
        </div>
    `,d}document.addEventListener("DOMContentLoaded",()=>{const s=document.getElementById("dropZone"),o=document.getElementById("fileInput"),i=document.querySelectorAll(".btn-city"),a=document.getElementById("rocDate"),n=document.getElementById("tpCategoryCol"),r=document.getElementById("tpGenderCol"),d=document.getElementById("tpAgeCol"),m=document.getElementById("tpLivingCol"),y=document.getElementById("tpCmsCol"),t=document.getElementById("ntGenderCol"),c=document.getElementById("ntAgeCol"),p=document.getElementById("ntLivingCol"),P=document.getElementById("ntCmsCol"),v=document.getElementById("previewSection"),l=document.getElementById("previewContainer"),O=document.getElementById("uploadBtn"),U=document.getElementById("statusMessage"),G=document.getElementById("fileInfo"),ee=document.getElementById("removeFile"),se=document.getElementById("caseListCols"),te=document.getElementById("serviceRegisterCols"),z=document.getElementById("cityInputGroup"),ae=document.getElementById("tpServiceCategoryGroup"),K=document.getElementById("headerRowIdx"),q=document.getElementById("districtCol"),X=document.getElementById("categoryCol"),H=document.getElementById("subsidyCol"),_=document.getElementById("serviceStatsGrid"),ne=document.getElementById("districtStats"),ie=document.getElementById("categoryStats"),ce=document.getElementById("subsidyStats"),Q=document.querySelectorAll(".btn-switch"),J=()=>{const e=E==="service"&&b==="台北";ae.classList.toggle("hidden",!e)},W="https://script.google.com/macros/s/AKfycbwfeDQPIplOO7e6uqxacl15bFoQ470bcGvufeXXFEQoAER7yjDBPrlI1WllKS8vrh2k/exec";let x=null,E="case",b="台北";z.classList.toggle("hidden",E==="case"),J(),Q.forEach(e=>{e.onclick=()=>{e.dataset.mode!==E&&(Q.forEach(f=>f.classList.remove("active")),e.classList.add("active"),E=e.dataset.mode,se.classList.toggle("hidden",E!=="case"),te.classList.toggle("hidden",E!=="service"),z.classList.toggle("hidden",E==="case"),J(),k())}}),i.forEach(e=>{e.onclick=()=>{e.dataset.city!==b&&(i.forEach(f=>f.classList.remove("active")),e.classList.add("active"),b=e.dataset.city,J(),k(),M())}}),document.querySelectorAll('input[name="tpServiceCategory"]').forEach(e=>{e.addEventListener("change",()=>{k()})}),s.onclick=()=>o.click(),s.ondragover=e=>{e.preventDefault(),s.classList.add("dragover")},s.ondragleave=()=>s.classList.remove("dragover"),s.ondrop=e=>{e.preventDefault(),s.classList.remove("dragover"),e.dataTransfer.files.length&&Z(e.dataTransfer.files[0])},o.onchange=e=>{e.target.files.length&&Z(e.target.files[0])},ee.onclick=e=>{e.stopPropagation(),k()};function Z(e){if(!e.name.match(/\.(xlsx|xls)$/)){j("不支援的檔案格式","error");return}x=e,F()}function F(){if(!x)return;const e=new FileReader;e.onload=f=>{try{const $=new Uint8Array(f.target.result),I=XLSX.read($,{type:"array"});if(E==="case"){const{taipeiSheetName:S,newTaipeiSheetName:C}=ve(I),g={taipei:S?XLSX.utils.sheet_to_json(I.Sheets[S],{header:1}):null,newTaipei:C?XLSX.utils.sheet_to_json(I.Sheets[C],{header:1}):null};oe(g,x.name,S,C)}else{const S=I.SheetNames[0],C=I.Sheets[S],g=XLSX.utils.sheet_to_json(C,{header:1});re(g,x.name)}}catch($){console.error($),j("讀取檔案失敗: "+$.message,"error")}},e.readAsArrayBuffer(x)}function k(){o.value="",x=null,G.classList.add("hidden"),document.querySelector(".upload-content").classList.remove("hidden"),v.classList.add("hidden"),O.disabled=!0,j("請選擇縣市、輸入月份並上傳檔案","")}function oe(e,f,$,I){l.innerHTML="",_&&_.classList.add("hidden"),l.classList.remove("hidden");const S=[],C=a.value;if(e.newTaipei){const g=me(e.newTaipei,{gender:t.value,age:c.value,living:p.value,cms:P.value}),h=`新北${C}`;l.appendChild(fe(g,h)),S.push({sheetName:h,rows:ge(g)})}if(e.taipei){const g={category:n.value,gender:r.value,age:d.value,living:m.value,cms:y.value},h=D(e.taipei,g,"老福"),T=D(e.taipei,g,"老福",{min:65}),R=D(e.taipei,g,"老福",{min:50,max:64}),L=D(e.taipei,g,"身障"),u=`台北老福${C}`,B=`台北身障${C}`;l.appendChild(pe(h,T,R,L,u,B));const N=ue(h,T,R,L);S.push({sheetName:u,rows:N.elderlyRows}),S.push({sheetName:B,rows:N.disabledRows})}O.dataset.outputPayload=JSON.stringify(S),V(f)}function re(e,f){const $=parseInt(K.value||"5"),I=w(q.value||"U"),S=w(X.value||"G"),C=w(H.value||"O"),g={},h={},T={};for(let R=$;R<e.length;R++){const L=e[R];if(I!==-1&&L[I]!==void 0){const u=L[I].toString().trim();u&&(g[u]=(g[u]||0)+1)}if(S!==-1&&L[S]!==void 0){const u=L[S].toString().trim();u&&(h[u]=(h[u]||0)+1)}if(C!==-1&&L[C]!==void 0){const u=L[C].toString().trim();u&&(T[u]=(T[u]||0)+1)}}de(g,h,T),V(f)}function V(e){G.classList.remove("hidden"),G.querySelector(".file-name").textContent=e,document.querySelector(".upload-content").classList.add("hidden"),v.classList.remove("hidden"),j("檔案已就緒，請確認統計結果","success"),M()}function de(e,f,$){l.classList.add("hidden"),_.classList.remove("hidden");const C=b==="台北"?["松山區","信義區","大安區","中山區","中正區","大同區","萬華區","文山區","南港區","內湖區","士林區","北投區"]:["三重區","蘆洲區","五股區","板橋區","土城區","新莊區","中和區","永和區","樹林區","泰山區","新店區"],g=Object.keys(e).sort((u,B)=>{const N=C.indexOf(u),A=C.indexOf(B);return N!==-1&&A!==-1?N-A:N!==-1?-1:A!==-1?1:u.localeCompare(B,"zh-hant")});ne.innerHTML=g.map(u=>`<div class="stat-item"><span>${u}</span> <span class="stat-value">${e[u]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const h=Object.keys(f).sort((u,B)=>u.localeCompare(B,"zh-hant"));ie.innerHTML=h.map(u=>`<div class="stat-item"><span>${u}</span> <span class="stat-value">${f[u]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const T=["100%","95%","84%"],R=Object.keys($).sort((u,B)=>{const N=T.indexOf(u),A=T.indexOf(B);return N!==-1&&A!==-1?N-A:N!==-1?-1:A!==-1?1:u.localeCompare(B)});ce.innerHTML=R.map(u=>`<div class="stat-item"><span>${u}</span> <span class="stat-value">${$[u]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const L=le(e,g,f,h,$,R);O.dataset.outputRows=JSON.stringify(L)}function le(e,f,$,I,S,C){const g=[["【行政區統計】",""]];return f.forEach(h=>g.push([h,e[h]])),g.push(["",""],["【服務項目類別】",""]),I.forEach(h=>g.push([h,$[h]])),g.push(["",""],["【補助比率】",""]),C.forEach(h=>g.push([h,S[h]])),g}function M(){const e=E==="case"||b!=="",f=a.value.length>=2,$=!!x;let I=!1;if(E==="case"){const S=n.value.trim()&&r.value.trim()&&d.value.trim()&&m.value.trim()&&y.value.trim(),C=t.value.trim()&&c.value.trim()&&p.value.trim()&&P.value.trim();I=S&&C}else I=q.value.trim()&&X.value.trim()&&H.value.trim();O.disabled=!(e&&f&&I&&$)}[a,n,r,d,m,y,t,c,p,P,q,X,H,K].forEach(e=>{e.addEventListener("input",()=>{M(),x&&F()}),e.addEventListener("change",()=>{M(),x&&F()})}),O.onclick=async()=>{if(E==="service"&&!b){j("請選擇縣市","error");return}if(a.value.length<2){j("請輸入月份","error");return}Y(!0),j("正在同步至 Google Sheet...","success");try{if(E==="case"){const e=JSON.parse(O.dataset.outputPayload||"[]");for(const f of e)await fetch(W,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify(f)})}else{const e=JSON.parse(O.dataset.outputRows||"[]");let f=b;b==="台北"&&(f=`台北${document.querySelector('input[name="tpServiceCategory"]:checked').value}`);const $=f+a.value;await fetch(W,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify({sheetName:$,rows:e})})}j("同步完成！請檢查您的 Google Sheet","success")}catch(e){console.error(e),j("同步失敗: "+e.message,"error")}finally{Y(!1)}};function Y(e){O.disabled=e,O.querySelector(".btn-text").classList.toggle("hidden",e),O.querySelector(".loader").classList.toggle("hidden",!e)}function j(e,f){U.textContent=e,U.className="status-message "+f}});
