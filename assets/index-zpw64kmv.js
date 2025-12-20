(function(){const o=document.createElement("link").relList;if(o&&o.supports&&o.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))a(n);new MutationObserver(n=>{for(const r of n)if(r.type==="childList")for(const l of r.addedNodes)l.tagName==="LINK"&&l.rel==="modulepreload"&&a(l)}).observe(document,{childList:!0,subtree:!0});function c(n){const r={};return n.integrity&&(r.integrity=n.integrity),n.referrerPolicy&&(r.referrerPolicy=n.referrerPolicy),n.crossOrigin==="use-credentials"?r.credentials="include":n.crossOrigin==="anonymous"?r.credentials="omit":r.credentials="same-origin",r}function a(n){if(n.ep)return;n.ep=!0;const r=c(n);fetch(n.href,r)}})();function ve(t){let o=null,c=null;return t.SheetNames.forEach(a=>{a.includes("台北")?o=a:a.includes("新北")&&(c=a)}),{taipeiSheetName:o,newTaipeiSheetName:c}}function E(t){if(!t)return-1;let o=0;const c=t.toUpperCase();for(let a=0;a<c.length;a++)o=o*26+c.charCodeAt(a)-64;return o-1}function ue(t,o){const c=o.headerRowIdx||5,a=E(o.district||""),n=E(o.category||""),r=E(o.subsidy||""),l={},p={},m={};for(let e=c;e<t.length;e++){const i=t[e];if(a!==-1&&i[a]!==void 0){const d=i[a].toString().trim();d&&(l[d]=(l[d]||0)+1)}if(n!==-1&&i[n]!==void 0){const d=i[n].toString().trim();d&&(p[d]=(p[d]||0)+1)}if(r!==-1&&i[r]!==void 0){const d=i[r].toString().trim();d&&(m[d]=(m[d]||0)+1)}}return{districtStats:l,categoryStats:p,subsidyStats:m}}function D(t,o,c,a=null){const n={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!t||t.length===0)return n;const r=E(o.category||""),l=E(o.gender||""),p=E(o.age||""),m=E(o.living||""),e=E(o.cms||"");for(let i=1;i<t.length;i++){const d=t[i];if((d[r]||"").toString().trim()!==c)continue;const u=parseInt(d[p]);if(!(a&&(isNaN(u)||a.min!==void 0&&u<a.min||a.max!==void 0&&u>a.max))){if(l!==-1&&d[l]!==void 0){const v=d[l].toString().trim();v==="男"?n.gender.男++:v==="女"&&n.gender.女++}if(isNaN(u)||(u<=49?n.age["<= 49"]++:u>=50&&u<=64?n.age["50 >= && <= 64"]++:u>=65&&u<=74?n.age["65 >= && <= 74"]++:u>=75&&u<=84?n.age["75 >= && <= 84"]++:u>=85&&n.age["85 >="]++),m!==-1&&d[m]!==void 0){let v=d[m].toString().trim();v&&(v==="其他類型居住狀況"&&(v="其他"),n.living[v]=(n.living[v]||0)+1)}if(e!==-1&&d[e]!==void 0){const v=d[e].toString().trim();v&&(n.cms[v]=(n.cms[v]||0)+1)}}}return n}function pe(t,o,c,a){const n=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],r=(e,i)=>{const d=[[`【${i}】`,""],["男",e.gender.男||0],["女",e.gender.女||0],["",""]];return d.push(["【年齡分佈】",""],["<= 49",e.age["<= 49"]||0],["50 >= && <= 64",e.age["50 >= && <= 64"]||0],["65 >= && <= 74",e.age["65 >= && <= 74"]||0],["75 >= && <= 84",e.age["75 >= && <= 84"]||0],["85 >=",e.age["85 >="]||0],["",""]),d},l=e=>{const i=["1","2","3","4","5","6","7","8"],d=Object.keys(e.cms).filter(v=>!i.includes(v)).sort(),j=i.concat(d),u=[["",""],["【CMS 等級】",""]];return j.forEach(v=>{e.cms[v]!==void 0&&u.push([v,e.cms[v]])}),u},p=()=>{let e=r(t,"性別統計");e.push(["(65歲以上 老人)",""]),n.forEach(u=>{e.push([u==="與家人或其他人同住"?"與家人住":u,o.living[u]||0])}),e.push(["",""],["(50~64 身障)",""]),["獨居","僅配偶2人","與家人住","與朋友住","其他"].forEach(u=>{let v=0;u==="僅配偶2人"?v=c.living["獨居(兩老)"]||0:u==="與家人住"?v=c.living.與家人或其他人同住||0:v=c.living[u]||0,e.push([u,v])}),e.push(["",""],["身心障礙者",""]);const d=Object.values(c.gender).reduce((u,v)=>u+v,0),j=Object.values(o.gender).reduce((u,v)=>u+v,0);return e.push(["50-64歲",d],["65歲以上",j]),e=e.concat(l(t)),e},m=()=>{let e=r(a,"性別統計");return e.push(["【居住狀況】",""]),n.forEach(i=>{a.living[i]!==void 0&&e.push([i,a.living[i]])}),e=e.concat(l(a)),e};return{elderlyRows:p(),disabledRows:m()}}function me(t,o,c,a,n,r){const l=["獨居","獨居(兩老)","與家人或其他人同住","與朋友住","其他"],p=Object.values(c.gender).reduce((i,d)=>i+d,0),m=Object.values(o.gender).reduce((i,d)=>i+d,0),e=document.createElement("div");return e.className="preview-group",e.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">台北老福</span>
            <span class="preview-group-badge">${n}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${t.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${t.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${t.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${t.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${t.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${t.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${t.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(t.cms).sort().map(i=>`<div class="stat-item"><span>${i}</span> <span class="stat-value">${t.cms[i]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>'}
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
    `,e}function ge(t,o){const c={gender:{男:0,女:0},age:{"<= 49":0,"50 >= && <= 64":0,"65 >= && <= 74":0,"75 >= && <= 84":0,"85 >=":0},living:{},cms:{}};if(!t||t.length===0)return c;const a=E(o.gender||""),n=E(o.age||""),r=E(o.living||""),l=E(o.cms||"");for(let p=1;p<t.length;p++){const m=t[p];if(a!==-1&&m[a]!==void 0){const e=m[a].toString().trim();e==="男"?c.gender.男++:e==="女"&&c.gender.女++}if(n!==-1&&m[n]!==void 0){const e=parseInt(m[n]);isNaN(e)||(e<=49?c.age["<= 49"]++:e>=50&&e<=64?c.age["50 >= && <= 64"]++:e>=65&&e<=74?c.age["65 >= && <= 74"]++:e>=75&&e<=84?c.age["75 >= && <= 84"]++:e>=85&&c.age["85 >="]++)}if(r!==-1&&m[r]!==void 0){let e=m[r].toString().trim();e&&(e==="其他類型居住狀況"&&(e="其他"),c.living[e]=(c.living[e]||0)+1)}if(l!==-1&&m[l]!==void 0){const e=m[l].toString().trim();e&&(c.cms[e]=(c.cms[e]||0)+1)}}return c}function fe(t){const o=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],c=["1","2","3","4","5","6","7","8"],a=Object.keys(t.cms).filter(l=>!c.includes(l)).sort(),n=c.concat(a),r=[["【性別統計】",""],["男",t.gender.男||0],["女",t.gender.女||0],["",""]];return r.push(["【年齡分佈】",""],["<= 49",t.age["<= 49"]||0],["50 >= && <= 64",t.age["50 >= && <= 64"]||0],["65 >= && <= 74",t.age["65 >= && <= 74"]||0],["75 >= && <= 84",t.age["75 >= && <= 84"]||0],["85 >=",t.age["85 >="]||0],["",""]),r.push(["【居住狀況】",""]),o.forEach(l=>{t.living[l]!==void 0&&r.push([l,t.living[l]])}),r.push(["",""],["【CMS 等級】",""]),n.forEach(l=>{t.cms[l]!==void 0&&r.push([l,t.cms[l]])}),r}function he(t,o){const c=["獨居","獨居(兩老)","與家人或其他人同住","與朋友同住","其他"],a=["1","2","3","4","5","6","7","8"],n=Object.keys(t.cms).filter(p=>!a.includes(p)).sort(),r=a.concat(n),l=document.createElement("div");return l.className="preview-group",l.innerHTML=`
        <div class="preview-group-header">
            <span class="preview-group-title">新北</span>
            <span class="preview-group-badge">${o}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${t.gender.男}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${t.gender.女}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${t.age["<= 49"]}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${t.age["50 >= && <= 64"]}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${t.age["65 >= && <= 74"]}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${t.age["75 >= && <= 84"]}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${t.age["85 >="]}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">居住狀況</div>
                <div class="stat-values">
                    ${c.filter(p=>t.living[p]!==void 0).map(p=>`<div class="stat-item"><span>${p}</span> <span class="stat-value">${t.living[p]}</span></div>`).join("")}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${r.filter(p=>t.cms[p]!==void 0).map(p=>`<div class="stat-item"><span>${p}</span> <span class="stat-value">${t.cms[p]}</span></div>`).join("")}
                </div>
            </div>
        </div>
    `,l}document.addEventListener("DOMContentLoaded",()=>{const t=document.getElementById("dropZone"),o=document.getElementById("fileInput"),c=document.querySelectorAll(".btn-city"),a=document.getElementById("rocDate"),n=document.getElementById("tpCategoryCol"),r=document.getElementById("tpGenderCol"),l=document.getElementById("tpAgeCol"),p=document.getElementById("tpLivingCol"),m=document.getElementById("tpCmsCol"),e=document.getElementById("ntGenderCol"),i=document.getElementById("ntAgeCol"),d=document.getElementById("ntLivingCol"),j=document.getElementById("ntCmsCol"),u=document.getElementById("previewSection"),v=document.getElementById("previewContainer"),L=document.getElementById("uploadBtn"),U=document.getElementById("statusMessage"),G=document.getElementById("fileInfo"),ee=document.getElementById("removeFile"),se=document.getElementById("caseListCols"),te=document.getElementById("serviceRegisterCols"),z=document.getElementById("cityInputGroup"),ae=document.getElementById("tpServiceCategoryGroup"),K=document.getElementById("headerRowIdx"),q=document.getElementById("districtCol"),X=document.getElementById("categoryCol"),H=document.getElementById("subsidyCol"),_=document.getElementById("serviceStatsGrid"),ne=document.getElementById("districtStats"),ie=document.getElementById("categoryStats"),ce=document.getElementById("subsidyStats"),Q=document.querySelectorAll(".btn-switch"),J=()=>{const s=w==="service"&&R==="台北";ae.classList.toggle("hidden",!s)},W="https://script.google.com/macros/s/AKfycbwfeDQPIplOO7e6uqxacl15bFoQ470bcGvufeXXFEQoAER7yjDBPrlI1WllKS8vrh2k/exec";let N=null,w="case",R="台北";z.classList.toggle("hidden",w==="case"),J(),Q.forEach(s=>{s.onclick=()=>{s.dataset.mode!==w&&(Q.forEach(g=>g.classList.remove("active")),s.classList.add("active"),w=s.dataset.mode,se.classList.toggle("hidden",w!=="case"),te.classList.toggle("hidden",w!=="service"),z.classList.toggle("hidden",w==="case"),J(),b())}}),c.forEach(s=>{s.onclick=()=>{s.dataset.city!==R&&(c.forEach(g=>g.classList.remove("active")),s.classList.add("active"),R=s.dataset.city,J(),b(),k())}}),document.querySelectorAll('input[name="tpServiceCategory"]').forEach(s=>{s.addEventListener("change",()=>{b()})}),t.onclick=()=>o.click(),t.ondragover=s=>{s.preventDefault(),t.classList.add("dragover")},t.ondragleave=()=>t.classList.remove("dragover"),t.ondrop=s=>{s.preventDefault(),t.classList.remove("dragover"),s.dataTransfer.files.length&&Z(s.dataTransfer.files[0])},o.onchange=s=>{s.target.files.length&&Z(s.target.files[0])},ee.onclick=s=>{s.stopPropagation(),b()};function Z(s){if(!s.name.match(/\.(xlsx|xls)$/)){x("不支援的檔案格式","error");return}N=s,F()}function F(){if(!N)return;const s=new FileReader;s.onload=g=>{try{const C=new Uint8Array(g.target.result),$=XLSX.read(C,{type:"array"});if(w==="case"){const{taipeiSheetName:S,newTaipeiSheetName:y}=ve($),f={taipei:S?XLSX.utils.sheet_to_json($.Sheets[S],{header:1}):null,newTaipei:y?XLSX.utils.sheet_to_json($.Sheets[y],{header:1}):null};oe(f,N.name,S,y)}else{const S=$.SheetNames[0],y=$.Sheets[S],f=XLSX.utils.sheet_to_json(y,{header:1});de(f,N.name)}}catch(C){console.error(C),x("讀取檔案失敗: "+C.message,"error")}},s.readAsArrayBuffer(N)}function b(){o.value="",N=null,G.classList.add("hidden"),document.querySelector(".upload-content").classList.remove("hidden"),u.classList.add("hidden"),L.disabled=!0,x("請選擇縣市、輸入月份並上傳檔案","")}function oe(s,g,C,$){v.innerHTML="",_&&_.classList.add("hidden"),v.classList.remove("hidden");const S=[],y=a.value;if(s.newTaipei){const f=ge(s.newTaipei,{gender:e.value,age:i.value,living:d.value,cms:j.value}),I=`新北${y}`;v.appendChild(he(f,I)),S.push({sheetName:I,rows:fe(f)})}if(s.taipei){const f={category:n.value,gender:r.value,age:l.value,living:p.value,cms:m.value},I=D(s.taipei,f,"老福"),P=D(s.taipei,f,"老福",{min:65}),A=D(s.taipei,f,"老福",{min:50,max:64}),M=D(s.taipei,f,"身障"),h=`台北老福${y}`,O=`台北身障${y}`;v.appendChild(me(I,P,A,M,h,O));const B=pe(I,P,A,M);S.push({sheetName:h,rows:B.elderlyRows}),S.push({sheetName:O,rows:B.disabledRows})}L.dataset.outputPayload=JSON.stringify(S),V(g)}function de(s,g){const C={headerRowIdx:parseInt(K.value||"5"),district:q.value||"U",category:X.value||"G",subsidy:H.value||"O"},{districtStats:$,categoryStats:S,subsidyStats:y}=ue(s,C);re($,S,y),V(g)}function V(s){G.classList.remove("hidden"),G.querySelector(".file-name").textContent=s,document.querySelector(".upload-content").classList.add("hidden"),u.classList.remove("hidden"),x("檔案已就緒，請確認統計結果","success"),k()}function re(s,g,C){v.classList.add("hidden"),_.classList.remove("hidden");const y=R==="台北"?["松山區","信義區","大安區","中山區","中正區","大同區","萬華區","文山區","南港區","內湖區","士林區","北投區"]:["三重區","蘆洲區","五股區","板橋區","土城區","新莊區","中和區","永和區","樹林區","泰山區","新店區"],f=Object.keys(s).sort((h,O)=>{const B=y.indexOf(h),T=y.indexOf(O);return B!==-1&&T!==-1?B-T:B!==-1?-1:T!==-1?1:h.localeCompare(O,"zh-hant")});ne.innerHTML=f.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${s[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const I=Object.keys(g).sort((h,O)=>h.localeCompare(O,"zh-hant"));ie.innerHTML=I.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${g[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const P=["100%","95%","84%"],A=Object.keys(C).sort((h,O)=>{const B=P.indexOf(h),T=P.indexOf(O);return B!==-1&&T!==-1?B-T:B!==-1?-1:T!==-1?1:h.localeCompare(O)});ce.innerHTML=A.map(h=>`<div class="stat-item"><span>${h}</span> <span class="stat-value">${C[h]}</span></div>`).join("")||'<div class="stat-item"><span>無資料</span></div>';const M=le(s,f,g,I,C,A);L.dataset.outputRows=JSON.stringify(M)}function le(s,g,C,$,S,y){const f=[["【行政區統計】",""]];return g.forEach(I=>f.push([I,s[I]])),f.push(["",""],["【服務項目類別】",""]),$.forEach(I=>f.push([I,C[I]])),f.push(["",""],["【補助比率】",""]),y.forEach(I=>f.push([I,S[I]])),f}function k(){const s=w==="case"||R!=="",g=a.value.length>=2,C=!!N;let $=!1;if(w==="case"){const S=n.value.trim()&&r.value.trim()&&l.value.trim()&&p.value.trim()&&m.value.trim(),y=e.value.trim()&&i.value.trim()&&d.value.trim()&&j.value.trim();$=S&&y}else $=q.value.trim()&&X.value.trim()&&H.value.trim();L.disabled=!(s&&g&&$&&C)}[a,n,r,l,p,m,e,i,d,j,q,X,H,K].forEach(s=>{s.addEventListener("input",()=>{k(),N&&F()}),s.addEventListener("change",()=>{k(),N&&F()})}),L.onclick=async()=>{if(w==="service"&&!R){x("請選擇縣市","error");return}if(a.value.length<2){x("請輸入月份","error");return}Y(!0),x("正在同步至 Google Sheet...","success");try{if(w==="case"){const s=JSON.parse(L.dataset.outputPayload||"[]");for(const g of s)await fetch(W,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify(g)})}else{const s=JSON.parse(L.dataset.outputRows||"[]");let g=R;R==="台北"&&(g=`台北${document.querySelector('input[name="tpServiceCategory"]:checked').value}`);const C=g+a.value;await fetch(W,{method:"POST",mode:"no-cors",cache:"no-cache",headers:{"Content-Type":"application/json"},body:JSON.stringify({sheetName:C,rows:s})})}b(),w==="case"&&(a.value=""),x("同步完成！請檢查您的 Google Sheet","success")}catch(s){console.error(s),x("同步失敗: "+s.message,"error")}finally{Y(!1)}};function Y(s){L.disabled=s,L.querySelector(".btn-text").classList.toggle("hidden",s),L.querySelector(".loader").classList.toggle("hidden",!s)}function x(s,g){U.textContent=s,U.className="status-message "+g}});
