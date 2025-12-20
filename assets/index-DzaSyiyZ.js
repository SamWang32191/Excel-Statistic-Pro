(function(){const o=document.createElement("link").relList;if(o&&o.supports&&o.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))a(n);new MutationObserver(n=>{for(const r of n)if(r.type==="childList")for(const l of r.addedNodes)l.tagName==="LINK"&&l.rel==="modulepreload"&&a(l)}).observe(document,{childList:!0,subtree:!0});function c(n){const r={};return n.integrity&&(r.integrity=n.integrity),n.referrerPolicy&&(r.referrerPolicy=n.referrerPolicy),n.crossOrigin==="use-credentials"?r.credentials="include":n.crossOrigin==="anonymous"?r.credentials="omit":r.credentials="same-origin",r}function a(n){if(n.ep)return;n.ep=!0;const r=c(n);fetch(n.href,r)}})();function ve(e){let o=null,c=null;return e.SheetNames.forEach(a=>{a.includes("台北")?o=a:a.includes("新北")&&(c=a)}),{taipeiSheetName:o,newTaipeiSheetName:c}}function $(e){if(!e)return-1;let o=0;const c=e.toUpperCase();for(let a=0;a<c.length;a++)o=o*26+c.charCodeAt(a)-64;return o-1}function pe(e,o){const c=o.headerRowIdx||5,a=$(o.district||""),n=$(o.category||""),r=$(o.subsidy||""),l={},p={},m={};for(let t=c;t<e.length;t++){const i=e[t];if(a!==-1&&i[a]!==void 0){const d=i[a].toString().trim();d&&(l[d]=(l[d]||0)+1)}if(n!==-1&&i[n]!==void 0){const d=i[n].toString().trim();d&&(p[d]=(p[d]||0)+1)}if(r!==-1&&i[r]!==void 0){const d=i[r].toString().trim();d&&(m[d]=(m[d]||0)+1)}}return{districtStats:l,categoryStats:p,subsidyStats:m}}function D(e,o,c,a=null){const n={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!e||e.length===0)return n;const r=$(o.category||""),l=$(o.gender||""),p=$(o.age||""),m=$(o.living||""),t=$(o.cms||"");for(let i=1;i<e.length;i++){const d=e[i];if((d[r]||"").toString().trim()!==c)continue;const v=parseInt(d[p]);if(!(a&&(isNaN(v)||a.min!==void 0&&v<a.min||a.max!==void 0&&v>a.max))){if(l!==-1&&d[l]!==void 0){const u=d[l].toString().trim();u==="男"?n.gender.男++:u==="女"&&n.gender.女++}if(isNaN(v)||(v<=49?n.age["<= 49"]++:v>=50&&v<=64?n.age["50 >= && <= 64"]++:v>=65&&v<=74?n.age["65 >= && <= 74"]++:v>=75&&v<=84?n.age["75 >= && <= 84"]++:v>=85&&n.age["85 >="]++),m!==-1&&d[m]!==void 0){let u=d[m].toString().trim();u&&(u==="其他類型居住狀況"&&(u="其他"),n.living[u]=(n.living[u]||0)+1)}if(t!==-1&&d[t]!==void 0){const u=d[t].toString().trim();u&&(n.cms[u]=(n.cms[u]||0)+1)}}}return n}function me(e,o,c,a){const n=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],r=(t,i)=>{const d=[[`【${i}】`,""],["男",t.gender.男||0],["女",t.gender.女||0],["",""]];return d.push(["【年齡分佈】",""],["<= 49",t.age["<= 49"]||0],["50 >= && <= 64",t.age["50 >= && <= 64"]||0],["65 >= && <= 74",t.age["65 >= && <= 74"]||0],["75 >= && <= 84",t.age["75 >= && <= 84"]||0],["85 >=",t.age["85 >="]||0],["",""]),d},l=t=>{const i=["1","2","3","4","5","6","7","8"],d=Object.keys(t.cms).filter(u=>!i.includes(u)).sort(),b=i.concat(d),v=[["",""],["【CMS 等級】",""]];return b.forEach(u=>{t.cms[u]!==void 0&&v.push([u,t.cms[u]])}),v},p=()=>{let t=r(e,"性別統計");t.push(["(65歲以上 老人)",""]),n.forEach(v=>{t.push([v==="與家人或其他人同住"?"與家人住":v,o.living[v]||0])}),t.push(["",""],["(50~64 身障)",""]),["獨居","僅配偶2人","與家人住","與朋友住","其他"].forEach(v=>{let u=0;v==="僅配偶2人"?u=c.living["獨居(兩老)"]||0:v==="與家人住"?u=c.living.與家人或其他人同住||0:u=c.living[v]||0,t.push([v,u])}),t.push(["",""],["身心障礙者",""]);const d=Object.values(c.gender).reduce((v,u)=>v+u,0),b=Object.values(o.gender).reduce((v,u)=>v+u,0);return t.push(["50-64歲",d],["65歲以上",b]),t=t.concat(l(e)),t},m=()=>{let t=r(a,"性別統計");return t.push(["【居住狀況】",""]),n.forEach(i=>{a.living[i]!==void 0&&t.push([i,a.living[i]])}),t=t.concat(l(a)),t};return{elderlyRows:p(),disabledRows:m()}}function ge(e,o,c,a,n,r){const l=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],p=Object.values(c.gender).reduce((i,d)=>i+d,0),m=Object.values(o.gender).reduce((i,d)=>i+d,0),t=document.createElement("div");return t.className="preview-group",t.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">台北老福</span>
            <span class="preview-group-badge">${n}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${e.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${e.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${e.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${e.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${e.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${e.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${e.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(e.cms).sort().map(i=>`<div class="stat-item"><span>${i}</span> <span class="stat-value">${e.cms[i]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
        <div class="stats-grid triples mt-4">
            <div class="stat-card">
                <div class="stat-label">(65歲以上 老人)</div>
                <div class="stat-values">
                    ${l.map(i=>`<div class="stat-item"><span>${i==="與家人或其他人同住"?"與家人住":i}</span> <span class="stat-value">${o.living[i]||0}</span></div>`).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">(50~64 身障)</div>
                <div class="stat-values">
                    ${["獨居","僅配偶2人","與家人住","與朋友住","其他"].map(i=>{let d=0;return i==="僅配偶2人"?d=c.living["獨居(兩老)"]||0:i==="與家人住"?d=c.living.與家人或其他人同住||0:d=c.living[i]||0,`<div class="stat-item"><span>${i}</span> <span class="stat-value">${d}</span></div>`}).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">身心障礙者</div>
                <div class="stat-values">
                    <div class="stat-item"><span>50-64歲</span> <span class="stat-value">${p}</span></div>
                    <div class="stat-item"><span>65歲以上</span> <span class="stat-value">${m}</span></div>
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
                    ${l.filter(i=>a.living[i]!==void 0).map(i=>`<div class="stat-item"><span>${i}</span> <span class="stat-value">${a.living[i]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(a.cms).sort().map(i=>`<div class="stat-item"><span>${i}</span> <span class="stat-value">${a.cms[i]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
    `,t}function fe(e,o){const c={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!e||e.length===0)return c;const a=$(o.gender||""),n=$(o.age||""),r=$(o.living||""),l=$(o.cms||"");for(let p=1;p<e.length;p++){const m=e[p];if(a!==-1&&m[a]!==void 0){const t=m[a].toString().trim();t==="男"?c.gender.男++:t==="女"&&c.gender.女++}if(n!==-1&&m[n]!==void 0){const t=parseInt(m[n]);isNaN(t)||(t<=49?c.age["<= 49"]++:t>=50&&t<=64?c.age["50 >= && <= 64"]++:t>=65&&t<=74?c.age["65 >= && <= 74"]++:t>=75&&t<=84?c.age["75 >= && <= 84"]++:t>=85&&c.age["85 >="]++)}if(r!==-1&&m[r]!==void 0){let t=m[r].toString().trim();t&&(t==="其他類型居住狀況"&&(t="其他"),c.living[t]=(c.living[t]||0)+1)}if(l!==-1&&m[l]!==void 0){const t=m[l].toString().trim();t&&(c.cms[t]=(c.cms[t]||0)+1)}}return c}function he(e){const o=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],c=["1","2","3","4","5","6","7","8"],a=Object.keys(e.cms).filter(l=>!c.includes(l)).sort(),n=c.concat(a),r=[["【性別統計】",""],["男",e.gender.男||0],["女",e.gender.女||0],["",""]];return r.push(["【年齡分佈】",""],["<= 49",e.age["<= 49"]||0],["50 >= && <= 64",e.age["50 >= && <= 64"]||0],["65 >= && <= 74",e.age["65 >= && <= 74"]||0],["75 >= && <= 84",e.age["75 >= && <= 84"]||0],["85 >=",e.age["85 >="]||0],["",""]),r.push(["【居住狀況】",""]),o.forEach(l=>{e.living[l]!==void 0&&r.push([l,e.living[l]])}),r.push(["",""],["【CMS 等級】",""]),n.forEach(l=>{e.cms[l]!==void 0&&r.push([l,e.cms[l]])}),r}function ye(e,o){const c=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],a=["1","2","3","4","5","6","7","8"],n=Object.keys(e.cms).filter(p=>!a.includes(p)).sort(),r=a.concat(n),l=document.createElement("div");return l.className="preview-group",l.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">新北</span>
            <span class="preview-group-badge">${o}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${e.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${e.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${e.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${e.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${e.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${e.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${e.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">居住狀況</div>
                <div class="stat-values">
                    ${c.filter(p=>e.living[p]!==void 0).map(p=>`<div class="stat-item"><span>${p}</span> <span class="stat-value">${e.living[p]}</span></div>`).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${r.filter(p=>e.cms[p]!==void 0).map(p=>`<div class="stat-item"><span>${p}</span> <span class="stat-value">${e.cms[p]}</span></div>`).join("")}
                </div>
            </div>
        </div>
    `,l}function Ce(){const e=document.getElementById("themeToggle");if(!e)return;const o=e.querySelectorAll(".btn-theme"),c=localStorage.getItem("theme")||"system";U(c),o.forEach(a=>{a.addEventListener("click",()=>{const n=a.getAttribute("data-theme");U(n),localStorage.setItem("theme",n)})}),window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change",()=>{(localStorage.getItem("theme")==="system"||!localStorage.getItem("theme"))&&U("system")})}function U(e){const o=document.documentElement;document.querySelectorAll(".btn-theme").forEach(r=>r.classList.remove("active"));let a=e;e==="system"&&(a=window.matchMedia("(prefers-color-scheme: dark)").matches?"dark":"light"),a==="dark"?o.removeAttribute("data-theme"):o.setAttribute("data-theme","light");const n=document.querySelector(`.btn-theme[data-theme="${e}"]`);n&&n.classList.add("active"),document.body.classList.remove("theme-dark","theme-light"),document.body.classList.add(`theme-${a}`)}document.addEventListener("DOMContentLoaded",()=>{Ce();const e=document.getElementById("dropZone"),o=document.getElementById("fileInput"),c=document.querySelectorAll(".btn-city"),a=document.getElementById("rocDate"),n=document.getElementById("tpCategoryCol"),r=document.getElementById("tpGenderCol"),l=document.getElementById("tpAgeCol"),p=document.getElementById("tpLivingCol"),m=document.getElementById("tpCmsCol"),t=document.getElementById("ntGenderCol"),i=document.getElementById("ntAgeCol"),d=document.getElementById("ntLivingCol"),b=document.getElementById("ntCmsCol"),v=document.getElementById("previewSection"),u=document.getElementById("previewContainer"),L=document.getElementById("uploadBtn"),z=document.getElementById("statusMessage"),G=document.getElementById("fileInfo"),te=document.getElementById("removeFile"),se=document.getElementById("caseListCols"),ae=document.getElementById("serviceRegisterCols"),K=document.getElementById("cityInputGroup"),ne=document.getElementById("tpServiceCategoryGroup"),Q=document.getElementById("headerRowIdx"),q=document.getElementById("districtCol"),X=document.getElementById("categoryCol"),H=document.getElementById("subsidyCol"),_=document.getElementById("serviceStatsGrid"),ie=document.getElementById("districtStats"),ce=document.getElementById("categoryStats"),oe=document.getElementById("subsidyStats"),W=document.querySelectorAll(".btn-switch"),J=()=>{const s=w==="service"&&j==="台北";ne.classList.toggle("hidden",!s)},Z="https://script.google.com/macros/s/AKfycbwfeDQPIplOO7e6uqxacl15bFoQ470bcGvufeXXFEQoAER7yjDBPrlI1WllKS8vrh2k/exec";let N=null,w="case",j="台北";K.classList.toggle("hidden",w==="case"),J(),W.forEach(s=>{s.onclick=()=>{s.dataset.mode!==w&&(W.forEach(g=>g.classList.remove("active")),s.classList.add("active"),w=s.dataset.mode,se.classList.toggle("hidden",w!=="case"),ae.classList.toggle("hidden",w!=="service"),K.classList.toggle("hidden",w==="case"),J(),R())}}),c.forEach(s=>{s.onclick=()=>{s.dataset.city!==j&&(c.forEach(g=>g.classList.remove("active")),s.classList.add("active"),j=s.dataset.city,J(),R(),P())}}),document.querySelectorAll('input[name="tpServiceCategory"]').forEach(s=>{s.addEventListener("change",()=>{R()})}),e.onclick=()=>o.click(),e.ondragover=s=>{s.preventDefault(),e.classList.add("dragover")},e.ondragleave=()=>e.classList.remove("dragover"),e.ondrop=s=>{s.preventDefault(),e.classList.remove("dragover"),s.dataTransfer.files.length&&V(s.dataTransfer.files[0])},o.onchange=s=>{s.target.files.length&&V(s.target.files[0])},te.onclick=s=>{s.stopPropagation(),R()};function V(s){if(!s.name.match(/\.(xlsx|xls)$/)){x("不支援的檔案格式","error");return}N=s,F()}function F(){if(!N)return;const s=new FileReader;s.onload=g=>{try{const C=new Uint8Array(g.target.result),E=XLSX.read(C,{type:"array"});if(w==="case"){const{taipeiSheetName:S,newTaipeiSheetName:y}=ve(E),f={taipei:S?XLSX.utils.sheet_to_json(E.Sheets[S],{header:1}):null,newTaipei:y?XLSX.utils.sheet_to_json(E.Sheets[y],{header:1}):null};re(f,N.name,S,y)}else{const S=E.SheetNames[0],y=E.Sheets[S],f=XLSX.utils.sheet_to_json(y,{header:1});de(f,N.name)}}catch(C){console.error(C),x("讀取檔案失敗: "+C.message,"error")}},s.readAsArrayBuffer(N)}function R(){o.value="",N=null,G.classList.add("hidden"),document.querySelector(".upload-content").classList.remove("hidden"),v.classList.add("hidden"),L.disabled=!0,x("請選擇縣市、輸入月份並上傳檔案","")}function re(s,g,C,E){u.innerHTML="",_&&_.classList.add("hidden"),u.classList.remove("hidden");const S=[],y=a.value;if(s.newTaipei){const f=fe(s.newTaipei,{gender:t.value,age:i.value,living:d.value,cms:b.value}),I=`新北${y}`;u.appendChild(ye(f,I)),S.push({sheetName:I,rows:he(f)})}if(s.taipei){const f={category:n.value,gender:r.value,age:l.value,living:p.value,cms:m.value},I=D(s.taipei,f,"老福"),k=D(s.taipei,f,"老福",{min:65}),A=D(s.taipei,f,"老福",{min:50,max:64}),M=D(s.taipei,f,"身障"),h=`台北老福${y}`,O=`台北身障${y}`;u.appendChild(ge(I,k,A,M,h,O));const B=me(I,k,A,M);S.push({sheetName:h,rows:B.elderlyRows}),S.push({sheetName:O,rows:B.disabledRows})}L.dataset.outputPayload=JSON.stringify(S),Y(g)}function de(s,g){const C={headerRowIdx:parseInt(Q.value||"5"),district:q.value||"U",category:X.value||"G",subsidy:H.value||"O"},{districtStats:E,categoryStats:S,subsidyStats:y}=pe(s,C);le(E,S,y),Y(g)}function Y(s){G.classList.remove("hidden"),G.querySelector(".file-name").textContent=s,document.querySelector(".upload-content").classList.add("hidden"),v.classList.remove("hidden"),x("檔案已就緒，請確認統計結果","success"),P()}function le(s,g,C){u.classList.add("hidden"),_.classList.remove("hidden");const y=j==="台北"?["松山區","信義區","大安區","中山區","中正區","大同區","萬華區","文山區","南港區","內湖區","士林區","北投區"]:["三重區","蘆洲區","五股區","板橋區","土城區","新莊區","中和區","永和區","樹林區","泰山區","新店區"],f=Object.keys(s).sort((h,O)=>{const B=y.indexOf(h),T=y.indexOf(O);return B!==-1&&T!==-1?B-T:B!==-1?-1:T!==-1?1:h.localeCompare(O,"zh-hant")});ie.innerHTML=f.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${s[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const I=Object.keys(g).sort((h,O)=>h.localeCompare(O,"zh-hant"));ce.innerHTML=I.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${g[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const k=["100%","95%","84%"],A=Object.keys(C).sort((h,O)=>{const B=k.indexOf(h),T=k.indexOf(O);return B!==-1&&T!==-1?B-T:B!==-1?-1:T!==-1?1:h.localeCompare(O)});oe.innerHTML=A.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${C[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const M=ue(s,f,g,I,C,A);L.dataset.outputRows=JSON.stringify(M)}function ue(s,g,C,E,S,y){const f=[["【行政區統計】",""]];return g.forEach(I=>f.push([I,s[I]])),f.push(["",""],["【服務項目類別】",""]),E.forEach(I=>f.push([I,C[I]])),f.push(["",""],["【補助比率】",""]),y.forEach(I=>f.push([I,S[I]])),f}function P(){const s=w==="case"||j!=="",g=a.value.length>=2,C=!!N;let E=!1;if(w==="case"){const S=n.value.trim()&&r.value.trim()&&l.value.trim()&&p.value.trim()&&m.value.trim(),y=t.value.trim()&&i.value.trim()&&d.value.trim()&&b.value.trim();E=S&&y}else E=q.value.trim()&&X.value.trim()&&H.value.trim();L.disabled=!(s&&g&&E&&C)}[a,n,r,l,p,m,t,i,d,b,q,X,H,Q].forEach(s=>{s.addEventListener("input",()=>{P(),N&&F()}),s.addEventListener("change",()=>{P(),N&&F()})}),L.onclick=async()=>{if(w==="service"&&!j){x("請選擇縣市","error");return}if(a.value.length<2){x("請輸入月份","error");return}ee(!0),x("正在同步至 Google Sheet...","success");try{if(w==="case"){const s=JSON.parse(L.dataset.outputPayload||"[]");for(const g of s)await fetch(Z,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify(g)})}else{const s=JSON.parse(L.dataset.outputRows||"[]");let g=j;j==="台北"&&(g=`台北${document.querySelector('input[name="tpServiceCategory"]:checked').value}`);const C=g+a.value;await fetch(Z,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify({sheetName:C,rows:s})})}R(),w==="case"&&(a.value=""),x("同步完成！請檢查您的 Google Sheet","success")}catch(s){console.error(s),x("同步失敗: "+s.message,"error")}finally{ee(!1)}};function ee(s){L.disabled=s,L.querySelector(".btn-text").classList.toggle("hidden",s),L.querySelector(".loader").classList.toggle("hidden",!s)}function x(s,g){z.textContent=s,z.className="status-message "+g}});
