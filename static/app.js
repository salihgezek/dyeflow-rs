
const TEXTS={EN:{tab_basic:"Basic Inputs",tab_machine:"Machine",tab_process:"Process",tab_steps:"Steps",tab_report:"Report",tab_dashboard:"Cost + Carbon",tab_projects:"Projects",tab_archive:"Archive",project_utility_inputs:"Project & Utility Inputs",project_name:"Project Name",company_name:"Company Name",heating_source:"Heating Source",heating_capacity:"Heating Capacity (Kcal/m3)",transfer_heat_loss:"Transfer Heat Loss (%)",natural_gas_unit_price:"Natural Gas Unit Price (€/Sm³)",water_type:"Water Type",water_unit_price:"Water Unit Price (€/m³)",waste_water_unit_price:"Waste Water Unit Price (€/m³)",electric_unit_price:"Electric Unit Price (€/kWh)",labour:"Labour",hourly_wage:"Hourly Wage",number_of_workers:"Number of Workers",number_of_machine:"Number of Machine",notes:"Notes",show_total_costs:"Show Total Costs",next_step:"Next Step",previous_step:"Previous Step",machine_information:"Machine Information",machine_name:"Machine Name",capacity_kg:"Capacity (Kg)",drain_time_min:"Drain Time (Min)",circulation_pump_power:"Circulation Pump Power (kW)",pump_ratio:"Pump Ratio",number_of_reel:"Number Of Reel",reel_power:"Reel Power (kW)",fan_power:"Fan Power (kW)",fan_ratio:"Fan Ratio",total_power_used:"Total Power Used In Machine",calculate_power:"Calculate Power",process:"Process",process_type:"Process Type",fabric_kg:"Fabric Kg",flote:"Flote",cost_currency:"Cost Currency",water_carry_over:"Water / Carry Over",carry_over:"Carry Over",fabric_status:"Fabric Status",sync_flote:"Transfer Flote and Water Amount to Steps",add_new_step:"+ ADD NEW STEP",duplicate_step:"Duplicate Step",calculate_draw:"Calculate and Draw Graph",summary:"Summary",select_report_image:"Select Report Image",copy_chemicals:"Copy Chemicals",presentation_package:"Presentation Package",export_powerpoint:"Export PowerPoint",export_excel:"Export Excel",export_pdf:"Export PDF",chart_title:"Dyeing Process Temperature Profile",cost_carbon_dashboard:"Cost + Carbon Dashboard",copy_dashboard_table:"Copy Dashboard Table",calculate_dashboard:"Calculate Dashboard",machine_archive:"Machine Archive",chemical_archive:"Chemical Archive",add_current_machine:"Add Current Machine to Archive",apply_selected_machine:"Apply Selected Machine",delete_selected:"Delete Selected",add_chemicals_from_project:"Add Project Chemicals to Archive",add_selected_chemical:"Add Selected Chemical to Current Step"},TR:{tab_basic:"Temel Girişler",tab_machine:"Makine",tab_process:"Proses",tab_steps:"Adımlar",tab_report:"Rapor",tab_dashboard:"Maliyet + Karbon",tab_projects:"Projeler",tab_archive:"Arşiv",project_utility_inputs:"Proje ve Yardımcı Girdiler",project_name:"Proje Adı",company_name:"Firma Adı",heating_source:"Isıtma Kaynağı",heating_capacity:"Isıtma Kapasitesi (Kcal/m3)",transfer_heat_loss:"Transfer Isı Kaybı (%)",natural_gas_unit_price:"Doğal Gaz Birim Fiyatı (€/Sm³)",water_type:"Su Tipi",water_unit_price:"Su Birim Fiyatı (€/m³)",waste_water_unit_price:"Atık Su Birim Fiyatı (€/m³)",electric_unit_price:"Elektrik Birim Fiyatı (€/kWh)",labour:"İşçilik",hourly_wage:"Saatlik Ücret",number_of_workers:"Çalışan Sayısı",number_of_machine:"Makine Sayısı",notes:"Notlar",show_total_costs:"Toplam Maliyetleri Göster",next_step:"Sonraki Adım",previous_step:"Önceki Adım",machine_information:"Makine Bilgileri",machine_name:"Makine Adı",capacity_kg:"Kapasite (Kg)",drain_time_min:"Boşaltma Süresi (Dk)",circulation_pump_power:"Sirkülasyon Pompa Gücü (kW)",pump_ratio:"Pompa Oranı",number_of_reel:"Haspel Sayısı",reel_power:"Haspel Gücü (kW)",fan_power:"Fan Gücü (kW)",fan_ratio:"Fan Oranı",total_power_used:"Makinede Kullanılan Toplam Güç",calculate_power:"Gücü Hesapla",process:"Proses",process_type:"Proses Tipi",fabric_kg:"Kumaş Kg",flote:"Flote",cost_currency:"Maliyet Para Birimi",water_carry_over:"Su / Taşıma",carry_over:"Taşıma",fabric_status:"Kumaş Durumu",sync_flote:"Flote ve Su Miktarını Step'lere Aktar",add_new_step:"+ YENİ ADIM EKLE",duplicate_step:"Step'i Çoğalt",calculate_draw:"Hesapla ve Grafiği Çiz",summary:"Özet",select_report_image:"Rapor Görseli Seç",copy_chemicals:"Kimyasalları Kopyala",presentation_package:"Sunum Paketi",export_powerpoint:"PowerPoint Aktar",export_excel:"Excel Aktar",export_pdf:"PDF Aktar",chart_title:"Boyama Proses Sıcaklık Profili",cost_carbon_dashboard:"Maliyet + Karbon Dashboard",copy_dashboard_table:"Dashboard Tablosunu Kopyala",calculate_dashboard:"Dashboard Hesapla",machine_archive:"Makine Arşivi",chemical_archive:"Kimyasal Arşivi",add_current_machine:"Mevcut Makineyi Arşive Ekle",apply_selected_machine:"Seçili Makineyi Uygula",delete_selected:"Seçileni Sil",add_chemicals_from_project:"Projedeki Kimyasalları Arşive Ekle",add_selected_chemical:"Seçili Kimyasalı Aktif Step'e Ekle"}};

let steps=[],last=null,selectedProject=null,selectedProjects=[],selectedMachineIndex=null,selectedChemicalIndex=null,machineArchive=[],chemicalArchive=[],graphZoom=1,graphPanX=0;
const $=id=>document.getElementById(id);

let authToken=localStorage.getItem("dyeflow_token")||"";
let currentUser=null;
let appInitialized=false;
function authHeaders(){return authToken?{"Authorization":"Bearer "+authToken}:{};}
function apiHeaders(extra={}){return Object.assign({"Content-Type":"application/json"},authHeaders(),extra);}
function showAuthModal(){ $("authModal")?.classList.remove("hidden"); }
function hideAuthModal(){ if(!currentUser){toast("Giriş yapılmadan uygulama kapatılamaz."); return;} $("authModal")?.classList.add("hidden"); }
function setForcedAccess(){
  const logged=!!currentUser;
  $("appShell")?.classList.toggle("hidden", !logged);
  if(logged){ $("authModal")?.classList.add("hidden"); } else { $("authModal")?.classList.remove("hidden"); }
  document.querySelector(".authClose")?.classList.toggle("hidden", !logged);
}
function initApp(){
  if(appInitialized) return;
  appInitialized=true;
  bindTabs();
  seedSteps();
  loadLocalArchives();
  renderSteps();
  renderArchives();
  applyLanguage();
  calculatePower();
  calculate();
}
function setAuthMode(mode){
  const reg=mode==="register";
  $("registerFields")?.classList.toggle("hidden",!reg);
  $("loginFields")?.classList.toggle("hidden",reg);
  $("loginTab")?.classList.toggle("active",!reg);
  $("registerTab")?.classList.toggle("active",reg);
  window.authMode=mode;
}
async function submitAuth(){
  const mode=window.authMode||"login";
  const payload=mode==="register"?{name:val("authName"),email:val("authEmail"),username:val("authUsername"),password:val("authPassword")}:{login:val("authLogin"),password:val("authPassword")};
  const res=await fetch(mode==="register"?"/api/auth/register":"/api/auth/login",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
  const data=await res.json();
  if(!res.ok){toast(data.error||"Login error");return;}
  authToken=data.token; currentUser=data.user; localStorage.setItem("dyeflow_token",authToken); updateAuthUi(); initApp(); refreshProjects(); toast("Logged in: "+(currentUser.username||"user"));
}
async function checkAuth(){
  if(!authToken){currentUser=null; updateAuthUi();return;}
  try{const res=await fetch("/api/auth/me",{headers:authHeaders()});const data=await res.json();if(data.authenticated){currentUser=data.user;}else{authToken="";localStorage.removeItem("dyeflow_token");currentUser=null;}}catch(e){currentUser=null;}
  updateAuthUi();
  if(currentUser) initApp();
}
function updateAuthUi(){
  setForcedAccess();
  const name=currentUser?(currentUser.name||currentUser.username):"Guest";
  if($("authStatus")) $("authStatus").innerText=currentUser?`User: ${name} (${currentUser.role||"user"})`:"Guest mode";
  $("loginBtn")?.classList.toggle("hidden",!!currentUser);
  $("logoutBtn")?.classList.toggle("hidden",!currentUser);
  $("adminTab")?.classList.toggle("hidden",!isAdmin());
  if(isAdmin()) adminRefreshUsers();
}
function logoutUser(){authToken="";currentUser=null;localStorage.removeItem("dyeflow_token");updateAuthUi();toast("Logged out");}
function requireLogin(){if(!authToken){showAuthModal();toast("Cloud save için önce login yapın. Admin tarafından verilen kullanıcı adı/şifre ile giriş yapmalısınız.");return false;}return true;}
function isAdmin(){return currentUser && currentUser.role==="admin";}
function boolText(v){return v?"Yes":"No";}
const num=id=>parseFloat($(id)?.value||0)||0;
const val=id=>$(id)?.value||"";
window.addEventListener("load",async()=>{
  setAuthMode('login');
  await checkAuth();
  if(currentUser){initApp(); refreshProjects();}
  else {setForcedAccess();}
});
function bindTabs(){document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>showTab(b.dataset.tab));$("reportImageInput").addEventListener("change",handleReportImage);}
function showTab(id){document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));$(id)?.classList.add("active");document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.tab===id));}
function applyLanguage(){const lang=$("language").value;document.querySelectorAll("[data-key]").forEach(el=>{const k=el.dataset.key;if(TEXTS[lang][k])el.textContent=TEXTS[lang][k];});}
function toast(m){const t=$("toast");t.innerText=m;t.style.opacity=1;setTimeout(()=>t.style.opacity=0,2200)}
function seedSteps(){steps=[{filling_time:4,beginning_temp:50,heating_slope:2,final_temp:90,dwelling_time:20,cooling_gradient:2,cooling_temp:70,overflow_time:0,amount_of_flote:3000,flote_ratio:6,bath_count:1,drain:false,chemicals:[{supplier:"Diğer",chemical:"Tanadye UCR",company:"",process:"",begin_c:50,final_c:50,dose_min:1,dose_time:1,circulation_time:0,amount:1,unit:"g/l",price:0}]}];}
function blankStep(){return{filling_time:0,beginning_temp:25,heating_slope:2,final_temp:60,dwelling_time:20,cooling_gradient:2,cooling_temp:40,overflow_time:0,amount_of_flote:num("fabric_kg")*num("flote"),flote_ratio:num("flote"),bath_count:1,drain:false,chemicals:[]};}
function blankChem(step){const t=Number(step.beginning_temp)||25;return{supplier:"Diğer",chemical:"",company:"",process:"",begin_c:t,final_c:t,dose_min:0,dose_time:0,circulation_time:0,amount:0,unit:"g/l",price:0};}
function addStep(){steps.push(blankStep());renderSteps();}
function duplicateSelectedOrLastStep(){steps.push(JSON.parse(JSON.stringify(steps[steps.length-1]||blankStep())));renderSteps();}
function duplicateStep(i){steps.splice(i+1,0,JSON.parse(JSON.stringify(steps[i])));renderSteps();}
function insertBeforeStep(i){steps.splice(i,0,JSON.parse(JSON.stringify(steps[i]||blankStep())));renderSteps();}
function deleteStep(i){if(steps.length<=1)return toast("At least one step must remain.");steps.splice(i,1);renderSteps();}
function addChemical(i){steps[i].chemicals.push(blankChem(steps[i]));renderSteps();}
function deleteChemical(i,j){steps[i].chemicals.splice(j,1);renderSteps();}
function updateStep(i,k,v){steps[i][k]=k==="drain"?v:(parseFloat(v)||0);}
function updateChem(i,j,k,v){if(["supplier","chemical","company","process","unit"].includes(k))steps[i].chemicals[j][k]=v;else steps[i].chemicals[j][k]=parseFloat(v)||0;}
function renderSteps(){const root=$("stepsContainer");root.innerHTML="";steps.forEach((s,i)=>{const box=document.createElement("div");box.className="stepBox";box.innerHTML=`<h3>Step ${i+1}</h3><div class="stepGrid"><label>Filling Time<input value="${s.filling_time}" onchange="updateStep(${i},'filling_time',this.value)"></label><label>Beginning Temp<input value="${s.beginning_temp}" onchange="updateStep(${i},'beginning_temp',this.value)"></label><label>Heating Slope<input value="${s.heating_slope}" onchange="updateStep(${i},'heating_slope',this.value)"></label><label>Final Temp<input value="${s.final_temp}" onchange="updateStep(${i},'final_temp',this.value)"></label><label>Dwelling Time<input value="${s.dwelling_time}" onchange="updateStep(${i},'dwelling_time',this.value)"></label><label>Cooling Gradient<input value="${s.cooling_gradient}" onchange="updateStep(${i},'cooling_gradient',this.value)"></label><label>Cooling Temp<input value="${s.cooling_temp}" onchange="updateStep(${i},'cooling_temp',this.value)"></label><label>Overflow Time<input value="${s.overflow_time}" onchange="updateStep(${i},'overflow_time',this.value)"></label><label>Amount Of Water<input value="${s.amount_of_flote}" onchange="updateStep(${i},'amount_of_flote',this.value)"></label><label>Flote Ratio<input value="${s.flote_ratio}" onchange="updateStep(${i},'flote_ratio',this.value)"></label><label><input type="checkbox" ${s.drain?"checked":""} onchange="updateStep(${i},'drain',this.checked)"> Drain</label></div><div class="chemRows">${s.chemicals.map((c,j)=>renderChemRow(c,i,j)).join("")}</div><button onclick="addChemical(${i})">+ Add Chemical</button><button onclick="insertBeforeStep(${i})">Insert Before Step</button><button onclick="duplicateStep(${i})">Duplicate Step</button><button onclick="deleteStep(${i})">Delete Step</button>`;root.appendChild(box);});}
function renderChemRow(c,i,j){return`<div class="chemGrid"><label>Supplier<input value="${c.supplier||""}" onchange="updateChem(${i},${j},'supplier',this.value)"></label><label>Chemical<input value="${c.chemical||""}" onchange="updateChem(${i},${j},'chemical',this.value)"></label><label>Company<input value="${c.company||""}" onchange="updateChem(${i},${j},'company',this.value)"></label><label>Process<input value="${c.process||""}" onchange="updateChem(${i},${j},'process',this.value)"></label><label>Begin °C<input value="${c.begin_c||0}" onchange="updateChem(${i},${j},'begin_c',this.value)"></label><label>Final °C<input value="${c.final_c||0}" onchange="updateChem(${i},${j},'final_c',this.value)"></label><label>Dose Min<input value="${c.dose_min||0}" onchange="updateChem(${i},${j},'dose_min',this.value)"></label><label>Dose Time<input value="${c.dose_time||0}" onchange="updateChem(${i},${j},'dose_time',this.value)"></label><label>Circulation Time<input value="${c.circulation_time||0}" onchange="updateChem(${i},${j},'circulation_time',this.value)"></label><label>Amount<input value="${c.amount||0}" onchange="updateChem(${i},${j},'amount',this.value)"></label><label>Unit<select onchange="updateChem(${i},${j},'unit',this.value)"><option ${c.unit==="g/l"?"selected":""}>g/l</option><option ${c.unit==="%"?"selected":""}>%</option></select></label><label>Price<input value="${c.price||0}" onchange="updateChem(${i},${j},'price',this.value)"></label><button onclick="deleteChemical(${i},${j})">Sil</button></div>`;}
function calculatePower(){const p=num("circulation_pump_power")*num("pump_ratio")+num("number_of_reel")*num("reel_power")+num("fan_power")*num("fan_ratio");$("total_power_used").value=p.toFixed(2);return p;}
function syncFlote(){steps.forEach(s=>{s.flote_ratio=num("flote");s.amount_of_flote=num("fabric_kg")*num("flote")});renderSteps();toast("Updated");}

async function adminRefreshUsers(){
  if(!isAdmin() || !$("adminUsersTable")) return;
  const r=await fetch("/api/admin/users",{headers:authHeaders()});
  const d=await r.json();
  if(!r.ok){toast(d.detail||d.error||"Admin user list error");return;}
  const rows=(d.users||[]).map(u=>`<tr>
    <td><b>${u.username||""}</b></td>
    <td>${u.email||""}</td>
    <td><input id="adm_name_${u.username}" value="${u.name||""}"></td>
    <td><select id="adm_role_${u.username}"><option value="user" ${u.role!=="admin"?"selected":""}>user</option><option value="admin" ${u.role==="admin"?"selected":""}>admin</option></select></td>
    <td><input id="adm_active_${u.username}" type="checkbox" ${u.is_active?"checked":""}></td>
    <td><input id="adm_save_${u.username}" type="checkbox" ${u.can_save?"checked":""}></td>
    <td><input id="adm_pass_${u.username}" type="password" placeholder="leave blank"></td>
    <td><button onclick="adminUpdateUser('${u.username}')">Save</button></td>
  </tr>`).join("");
  $("adminUsersTable").querySelector("tbody").innerHTML=rows;
}

async function adminCreateUser(){
  if(!isAdmin()) return toast("Admin only");
  const payload={
    name:val("adminName"), email:val("adminEmail"), username:val("adminUsername"), password:val("adminPassword"),
    role:val("adminRole")||"user", can_save:$("adminCanSave")?.checked, is_active:$("adminIsActive")?.checked
  };
  const r=await fetch("/api/admin/users",{method:"POST",headers:apiHeaders(),body:JSON.stringify(payload)});
  const d=await r.json();
  if(!r.ok){toast(d.error||d.detail||"Create user error");return;}
  ["adminName","adminEmail","adminUsername","adminPassword"].forEach(id=>{if($(id)) $(id).value="";});
  toast("User created: "+d.user.username);
  adminRefreshUsers();
}

async function adminUpdateUser(username){
  if(!isAdmin()) return toast("Admin only");
  const pass=$("adm_pass_"+username)?.value||"";
  const payload={
    name:$("adm_name_"+username)?.value||"",
    role:$("adm_role_"+username)?.value||"user",
    is_active:$("adm_active_"+username)?.checked,
    can_save:$("adm_save_"+username)?.checked
  };
  if(pass) payload.password=pass;
  const r=await fetch("/api/admin/users/"+encodeURIComponent(username),{method:"PATCH",headers:apiHeaders(),body:JSON.stringify(payload)});
  const d=await r.json();
  if(!r.ok){toast(d.error||d.detail||"Update user error");return;}
  toast("User updated: "+username);
  adminRefreshUsers();
}

function collectProject(){return{project_name:val("project_name"),company_name:val("company_name"),process_type:val("process_type"),fabric_kg:num("fabric_kg"),flote:num("flote"),cost_currency:document.querySelector("input[name=currency]:checked")?.value||"EUR",carry_over:num("carry_over"),fabric_status:document.querySelector("input[name=fabric_status]:checked")?.value||"Dry",notes:val("notes"),machine:{machine_name:val("machine_name"),capacity_kg:num("capacity_kg"),drain_time_min:num("drain_time_min"),circulation_pump_power:num("circulation_pump_power"),pump_ratio:num("pump_ratio"),number_of_reel:num("number_of_reel"),reel_power:num("reel_power"),fan_power:num("fan_power"),fan_ratio:num("fan_ratio")},utilities:{heating_source:val("heating_source"),heating_capacity:num("heating_capacity"),transfer_heat_loss:num("transfer_heat_loss"),natural_gas_unit_price:num("natural_gas_unit_price"),water_unit_price:num("water_unit_price"),waste_water_unit_price:num("waste_water_unit_price"),electric_unit_price:num("electric_unit_price"),hourly_wage:num("hourly_wage"),number_of_workers:num("number_of_workers"),number_of_machine:num("number_of_machine")},steps};}
function applyProject(d){$("project_name").value=d.project_name||"";$("company_name").value=d.company_name||"";$("process_type").value=d.process_type||"";$("fabric_kg").value=d.fabric_kg||0;$("flote").value=d.flote||0;$("carry_over").value=d.carry_over||0;$("notes").value=d.notes||"";if(d.machine)Object.entries(d.machine).forEach(([k,v])=>{if($(k))$(k).value=v;});if(d.utilities)Object.entries(d.utilities).forEach(([k,v])=>{if($(k))$(k).value=v;});steps=d.steps&&d.steps.length?d.steps:steps;renderSteps();calculatePower();}
async function calculate(){calculatePower();const res=await fetch("/api/calculate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(collectProject())});last=await res.json();drawChart(last);renderSummary(last);renderDashboard(last);renderChemicalCopy(last);renderArchiveChem(last);}
function calculateAndGoReport(){calculate().then(()=>showTab("report"));}
function drawChart(data){
 const chart=$("chart");if(!chart)return;
 const xs=data.x||[],ys=data.y||[];
 const vx=xs.filter(v=>v!==null&&v!==undefined),vy=ys.filter(v=>v!==null&&v!==undefined);
 if(!vx.length||!vy.length){chart.innerHTML="<div style='padding:20px'>No graph data.</div>";return;}
 const W=1280,H=470,ml=58,mr=26,mt=30,mb=42;
 const maxX=Math.max(...vx,10),minY=0,maxY=Math.max(110,Math.ceil((Math.max(...vy,100)+8)/10)*10);
 const X=t=>ml+(t/maxX)*(W-ml-mr);
 const Y=v=>mt+((maxY-v)/(maxY-minY))*(H-mt-mb);
 const safe=s=>String(s??"").replace(/[&<>]/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[m]));
 const N=(v,d=0)=>{v=parseFloat(v);return Number.isFinite(v)?v:d};

 function buildAnnotations(){
  let anns={bands:[],vlabels:[],stepLabels:[],drains:[]};
  let t=0,cur=null,letterIndex=0,letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const drainTime=N($("drain_time_min")?.value,5);
  steps.forEach((s,si)=>{
    let begin=N(s.beginning_temp,25);cur=begin;
    let stepStart=t;
    let fill=N(s.filling_time,0);
    anns.stepLabels.push({x:stepStart,label:`Filling\nStep ${si+1}`});
    if(fill>0){anns.bands.push({x1:t,x2:t+fill});t+=fill;anns.vlabels.push({x:t,y:begin,text:`${begin.toFixed(1)} °C`});}

    // Chemical timing logic v11: Dose Min is always based on
    // Step Start + Filling Time + Dose Min. Temperature ramps happen before
    // dosing; if the target temperature is reached early, wait until that
    // scheduled minute.
    let fillEndT=stepStart+fill;
    let groups={};
    (s.chemicals||[]).forEach(c=>{if(!["chemical","supplier","company"].some(k=>String(c[k]||"").trim()))return;let cb=N(c.begin_c,cur),dm=Math.max(N(c.dose_min,0),0),target=fillEndT+dm;let key=`${target.toFixed(3)}|${cb.toFixed(3)}|${dm.toFixed(3)}`;(groups[key]||(groups[key]=[])).push(c);});
    Object.keys(groups).sort((a,b)=>{let A=a.split('|').map(Number),B=b.split('|').map(Number);return A[0]-B[0]||A[1]-B[1];}).forEach(key=>{
      let [target,cb,dm]=key.split('|').map(Number),chems=groups[key];
      if(Math.abs(cb-cur)>0.01){let rate=cb>cur?Math.max(N(s.heating_slope,1),.1):Math.max(N(s.cooling_gradient,1),.1);t+=Math.abs(cb-cur)/rate;cur=cb;anns.vlabels.push({x:t,y:cur,text:`${rate.toFixed(1)} °C/min.`});}
      if(t<target){let wait=target-t;t=target;if(wait>.01)anns.vlabels.push({x:t,y:cur,text:`${dm.toFixed(1)} min.`});}
      chems.forEach(()=>letterIndex++);
      let doseTime=Math.max(...chems.map(c=>N(c.dose_time,0)),0);
      let endTemp=chems.length?N(chems[chems.length-1].final_c,cb):cb;
      if(doseTime>0){t+=doseTime;cur=endTemp;anns.vlabels.push({x:t,y:cur,text:`${doseTime.toFixed(1)} min.`});}
      let circ=Math.max(...chems.map(c=>N(c.circulation_time,0)),0);
      if(circ>0){t+=circ;anns.vlabels.push({x:t,y:cur,text:`${circ.toFixed(1)} min.`});}
    });
    let final=N(s.final_temp,cur);
    if(Math.abs(final-cur)>0.01){let rate=final>cur?Math.max(N(s.heating_slope,1),.1):Math.max(N(s.cooling_gradient,1),.1);t+=Math.abs(final-cur)/rate;cur=final;anns.vlabels.push({x:t,y:cur,text:`${final.toFixed(1)} °C`});}
    let dwell=N(s.dwelling_time,0);if(dwell>0){t+=dwell;anns.vlabels.push({x:t,y:cur,text:`${dwell.toFixed(1)} min.`});}
    let cool=N(s.cooling_temp,cur);if(cool<cur){let rate=Math.max(N(s.cooling_gradient,1),.1);t+=Math.abs(cur-cool)/rate;cur=cool;anns.vlabels.push({x:t,y:cur,text:`${rate.toFixed(1)} °C/min.`});}
    let overflow=N(s.overflow_time,0);if(overflow>0){t+=overflow;anns.vlabels.push({x:t,y:cur,text:`${overflow.toFixed(1)} min.`});}
    if(s.drain){anns.drains.push({x:t,y:cur});t+=drainTime;}
  });
  return anns;
 }
 const anns=buildAnnotations();

 let grid="";
 for(let v=0;v<=maxY;v+=20){let yy=Y(v);grid+=`<line x1="${ml}" y1="${yy}" x2="${W-mr}" y2="${yy}" class="gridH"/><text x="14" y="${yy+4}" class="axis">${v}</text>`;}
 for(let i=0;i<=10;i++){let xx=ml+i*(W-ml-mr)/10;let val=(i*maxX/10);grid+=`<line x1="${xx}" y1="${mt}" x2="${xx}" y2="${H-mb}" class="gridV"/><text x="${xx-10}" y="${H-13}" class="axis">${val.toFixed(0)}</text>`;}
 let bands=anns.bands.map(b=>`<rect x="${X(b.x1)}" y="${mt}" width="${Math.max(2,X(b.x2)-X(b.x1))}" height="${H-mt-mb}" class="fillBand"/>`).join("");
 let overflowBands=(data.events||[]).filter(e=>e.type==="overflow").map(e=>`<rect x="${X(e.x1)}" y="${mt}" width="${Math.max(2,X(e.x2)-X(e.x1))}" height="${H-mt-mb}" class="overflowBand"/><text x="${(X(e.x1)+X(e.x2))/2}" y="${H-30}" text-anchor="middle" class="overflowText">Overflow</text>`).join("");
 let stepLabels=anns.stepLabels.map(s=>`<text x="${X(s.x)+2}" y="${H-28}" class="smallNote">Filling</text><text x="${X(s.x)+2}" y="${H-15}" class="smallNote">Step ${safe(s.label.split('Step ')[1]||'')}</text>`).join("");
 let vlabels=anns.vlabels.map((a,idx)=>{let x=X(a.x),y=Math.max(mt+36,Math.min(H-mb-28,Y(a.y||50)+52));return `<line x1="${x}" y1="${mt}" x2="${x}" y2="${H-mb}" class="dashLine"/><text transform="translate(${x-3},${y}) rotate(-90)" class="rotLabel">${safe(a.text)}</text>`;}).join("");

 let seg="",curPts=[];
 for(let i=0;i<xs.length;i++){
  if(xs[i]===null||ys[i]===null){if(curPts.length){seg+=`<polyline class="temp" points="${curPts.join(" ")}"/>`;curPts=[];}}
  else curPts.push(`${X(xs[i])},${Y(ys[i])}`);
 }
 if(curPts.length)seg+=`<polyline class="temp" points="${curPts.join(" ")}"/>`;

 let ev="",used=[];
 function labelPos(x,y){let xx=x-7,yy=y-34;for(const p of used){if(Math.abs(xx-p[0])<30&&Math.abs(yy-p[1])<26)yy-=28;}used.push([xx,yy]);return[xx,yy];}
 (data.events||[]).forEach(e=>{
  if(e.type==="chemical_group"){
    let x=X(e.x),y=Y(e.y),label=(e.labels||[]).join(",");
    if((e.labels||[]).length>=4)label=(e.labels||[]).join("\n");
    if(e.dose_time>1){
      let x2=X(e.x+e.dose_time),y2=Y(e.y_end);
      let topY=Math.min(y,y2)-24;
      // Dosing triangle grows from low to high: narrow at start, wide at end.
      ev+=`<polygon points="${x},${y} ${x2},${y2} ${x2},${topY}" class="doseTri"/>`;
      ev+=`<text x="${(x+x2)/2}" y="${Math.min(topY,y2)-9}" class="chemLabel">${safe(label)}</text>`;
    }else{
      let p=labelPos(x,y);
      ev+=`<line x1="${p[0]+7}" y1="${p[1]+5}" x2="${x}" y2="${y-3}" class="chemArrow" marker-end="url(#blackArrow)"/>`;
      ev+=`<rect x="${p[0]-2}" y="${p[1]-14}" width="18" height="18" rx="2" class="chemBox"/><text x="${p[0]+3}" y="${p[1]}" class="chemBoxText">${safe(label)}</text>`;
    }
  }
 });
 (data.events||[]).filter(e=>e.type==="drain").forEach(e=>{let x=X(e.x),yStart=Y(e.y??50),yEnd=Math.min(H-mb-18,yStart+70);ev+=`<line x1="${x}" y1="${yStart}" x2="${x}" y2="${yEnd}" class="drainGuide"/><line x1="${x}" y1="${yStart}" x2="${x}" y2="${yEnd}" class="drainArrow" marker-end="url(#orangeArrow)"/><rect x="${x-20}" y="${yEnd+4}" width="40" height="16" rx="4" class="drainTag"/><text x="${x}" y="${yEnd+16}" text-anchor="middle" class="drainText">Drain</text>`;});

 const visibleW=W/graphZoom;
 const panLimit=Math.max(0,(W-visibleW));
 const viewX=Math.max(0,Math.min(panLimit,graphPanX));
 chart.innerHTML=`<div class="graphHint">Zoom: ${Math.round(graphZoom*100)}% • Pan: ${Math.round(viewX)} • chemical labels grouped by dosing minute</div><svg viewBox="${viewX} 0 ${visibleW} ${H}" width="100%" height="100%" preserveAspectRatio="xMidYMid meet"><defs><pattern id="hatch" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="8" stroke="#eef2f6" stroke-width="4"/></pattern><marker id="blackArrow" markerWidth="8" markerHeight="8" refX="4" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#111"/></marker><marker id="orangeArrow" markerWidth="8" markerHeight="8" refX="4" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#ff7a00"/></marker></defs><style>.gridH{stroke:#e6e6e6;stroke-width:1}.gridV{stroke:#d9d9d9;stroke-width:1;stroke-dasharray:3 3}.dashLine{stroke:#cfcfcf;stroke-width:1;stroke-dasharray:3 3}.axis{font:14px Segoe UI,Arial;fill:#222}.temp{fill:none;stroke:#e60012;stroke-width:3;stroke-linecap:square;stroke-linejoin:miter}.fillBand{fill:url(#hatch);opacity:.95}.overflowBand{fill:#dff3ff;opacity:.45}.overflowText{font:bold 11px Segoe UI,Arial;fill:#4f7182}.smallNote{font:bold 12px Segoe UI,Arial;fill:#8a929b}.rotLabel{font:bold 12px Segoe UI,Arial;fill:#8f8f8f}.chemArrow{stroke:#111;stroke-width:1.7}.chemBox{fill:#fff;stroke:#333;stroke-width:1.3}.chemBoxText{font:bold 12px Segoe UI,Arial;fill:#111}.doseTri{fill:none;stroke:#111;stroke-width:2.2}.drainGuide{stroke:#ff7a00;stroke-width:5;opacity:.16}.drainArrow{stroke:#ff7a00;stroke-width:2.2;stroke-dasharray:7 4}.drainTag{fill:#fff4e8;stroke:#ff7a00;stroke-width:1}.drainText{font:bold 11px Segoe UI,Arial;fill:#b65300}.title{font:bold 20px Segoe UI,Arial;fill:#002856}.labelAxis{font:bold 13px Segoe UI,Arial;fill:#111}</style><rect x="0" y="0" width="${W}" height="${H}" fill="#fff"/><text x="${W/2}" y="22" text-anchor="middle" class="title">Dyeing Process Temperature Profile</text><text transform="translate(10,${H/2+70}) rotate(-90)" class="labelAxis">Temperature (°C)</text><text x="${W/2}" y="${H-2}" text-anchor="middle" class="labelAxis">Time (min)</text>${bands}${overflowBands}${grid}${vlabels}${stepLabels}${seg}${ev}<rect x="${ml}" y="${mt}" width="${W-ml-mr}" height="${H-mt-mb}" fill="none" stroke="#111" stroke-width="1.2"/></svg>`;
 chart.querySelectorAll(".chemBoxText,.chemBox").forEach(el=>{el.style.cursor="help";});
}
function setGraphZoom(f){graphZoom=Math.min(3,Math.max(0.6,graphZoom*f)); if(last)drawChart(last);}
function panGraph(delta){graphPanX+=delta; if(last)drawChart(last);}
function resetGraphZoom(){graphZoom=1; graphPanX=0; if(last)drawChart(last);}
function openReportHtml(){calculate().then(()=>postDownload("/api/export/report-html","DyeFlow_RS_Report.html"));}

function dashVal(d, keys, fallback=0){
 for(const k of keys){ if(d[k]!==undefined && d[k]!==null && d[k]!=="") return d[k]; }
 return fallback;
}
function dashNum(d, keys, fallback=0){ return Number(dashVal(d, keys, fallback))||0; }
function renderSummary(data){
 const d=data.dashboard||{};
 const totalCost=dashVal(d,["Total Cost","Total Cost / batch (EUR)","Total Cost / batch (USD)","Total Cost / batch (TRY)"],0);
 $("summaryBox").innerHTML=`<div class="summaryGrid"><div class="summaryItem"><small>Total Time</small><b>${dashVal(d,["Total Time (min)"],data.total_time||0)} min</b></div><div class="summaryItem"><small>Total Chemicals</small><b>${(data.chemical_rows||[]).length}</b></div><div class="summaryItem"><small>Total Cost</small><b>${totalCost}</b></div></div>`;
}
function renderChemicalCopy(data){$("chemicalCopy").value=(data.chemical_legend||[]).join("\n");}
function renderDashboard(data){
 const d=data.dashboard||{};
 const heatingConsumptionKey=Object.keys(d).find(k=>k.startsWith("Heating Consumption"))||"Heating Consumption";
 const totalCostKey=Object.keys(d).find(k=>k.startsWith("Total Cost / batch"))||"Total Cost";
 const metricItems=[
  ["Total Cost / batch", dashVal(d,[totalCostKey,"Total Cost"],0)],
  ["Total Cost / kg", dashVal(d,["Total Cost / kg","Cost / kg"],0)],
  ["Total Time (min)", dashVal(d,["Total Time (min)"],data.total_time||0)],
  ["Electricity (kWh)", dashVal(d,["Electricity (kWh/batch)","Electricity (kWh)"],0)],
  ["Heating Energy (kcal)", dashVal(d,["Heating Energy (kcal/batch)","Heating Energy (kcal)"],0)],
  [heatingConsumptionKey, dashVal(d,[heatingConsumptionKey],0)],
  ["Heating Cost / batch", dashVal(d,["Heating Cost / batch","Heating Cost"],0)],
  ["Heating CO₂ kg/batch", dashVal(d,["Heating CO₂ (kg/batch)","Heating CO₂ (kg)"],0)],
  ["Total CO₂ kg/batch", dashVal(d,["Total CO₂ (kg/batch)","Total CO₂ (kg)"],0)],
  ["Total CO₂ g/kg", dashVal(d,["Total CO₂ / kg (g)","CO₂ / kg (g)"],0)]
 ];
 $("metrics").innerHTML=metricItems.map(([k,v])=>`<div class="card"><small>${k}</small><b>${v}</b></div>`).join("");
 const total=dashNum(d,[totalCostKey,"Total Cost"],0);
 const rows=[
  ["Chemical",dashNum(d,["Chemical Cost / batch","Chemical Cost"],0)],
  ["Electricity",dashNum(d,["Electricity Cost / batch","Electricity Cost"],0)],
  ["Heating",dashNum(d,["Heating Cost / batch","Heating Cost"],0)],
  ["Water",dashNum(d,["Water Cost / batch","Water Cost"],0)],
  ["Waste Water",dashNum(d,["Waste Water Cost / batch","Waste Water Cost"],0)],
  ["Labour",dashNum(d,["Labour Cost / batch","Labour Cost"],0)]
 ];
 $("costTable").querySelector("tbody").innerHTML=rows.map(r=>`<tr><td>${r[0]}</td><td>${r[1].toFixed(2)}</td><td>${num("fabric_kg")?(r[1]/num("fabric_kg")).toFixed(3):"-"}</td><td>${total?(r[1]/total*100).toFixed(1):"-"}</td></tr>`).join("");
 const fabric=num("fabric_kg");
 const ec=dashNum(d,["Electricity CO₂ (kg/batch)","Electricity CO₂ (kg)"],0), hc=dashNum(d,["Heating CO₂ (kg/batch)","Heating CO₂ (kg)"],0), tc=dashNum(d,["Total CO₂ (kg/batch)","Total CO₂ (kg)"],0);
 const carbonRows=[["Electricity CO₂",ec],["Heating CO₂",hc],["Total CO₂",tc]];
 $("carbonTable").querySelector("tbody").innerHTML=carbonRows.map(r=>`<tr><td>${r[0]}</td><td>${r[1].toFixed(2)}</td><td>${fabric?(r[1]*1000/fabric).toFixed(2):"-"}</td><td>${tc?(r[1]/tc*100).toFixed(1):"-"}</td></tr>`).join("");
 const maxCost=Math.max(...rows.map(r=>r[1]),1);
 const bars=rows.map(r=>{const pct=Math.round((r[1]/maxCost)*100);const share=total?(r[1]/total*100).toFixed(1):0;return `<div class="barRow"><b>${r[0]}</b><div class="barTrack"><div class="barFill" style="width:${pct}%"></div></div><span>${share}%</span></div>`;}).join("");
 const co2=tc; const perKg=dashNum(d,["Total CO₂ / kg (g)","CO₂ / kg (g)"],0); const gaugePct=Math.min(100,Math.max(5,perKg/10));
 const dv=$("dashboardVisuals"); if(dv){dv.innerHTML=`<div class="visualCard"><div class="visualTitle">Cost Breakdown</div>${bars}</div><div class="visualCard"><div class="visualTitle">CO₂ Intensity</div><div class="co2Gauge" style="--p:${gaugePct}%"><b>${perKg}</b></div><p style="text-align:center;color:#63748a;font-weight:800;margin:0">g CO₂ / kg • ${co2.toFixed(2)} kg/batch</p></div>`;}
}
function renderArchiveChem(data){if(!chemicalArchive.length){chemicalArchive=(data.chemical_rows||[]).map(r=>({supplier:r.supplier||"",chemical:r.chemical||"",amount:r.amount||0,unit:r.unit||"g/l",price:r.price||0}));saveLocalArchives();}renderArchives();}
function copyChemicals(){navigator.clipboard.writeText($("chemicalCopy").value);toast("Copied");}
function copyDashboard(){navigator.clipboard.writeText(Object.entries(last?.dashboard||{}).map(([k,v])=>`${k}\t${v}`).join("\n"));toast("Copied");}
async function saveProject(){if(!requireLogin())return;const r=await fetch("/api/save-project",{method:"POST",headers:apiHeaders(),body:JSON.stringify(collectProject())});const d=await r.json();if(!r.ok){toast(d.detail||d.error||"Save error");return;}toast("Saved: "+d.file);refreshProjects();}
async function refreshProjects(){if(!requireLogin())return;const r=await fetch("/api/projects",{headers:authHeaders()});const d=await r.json();if(!r.ok){toast(d.detail||d.error||"Project list error");return;}$("projectsTable").querySelector("tbody").innerHTML=(d.projects||[]).map(item=>{const p=typeof item==="string"?item:item.file;const up=typeof item==="string"?"-":(item.updated||"-");const pn=typeof item==="string"?"-":(item.project_name||"-");const co=typeof item==="string"?"-":(item.company_name||"-");return `<tr onclick="selectProject('${p}',this)"><td>${p}</td><td>${up}</td><td>${pn}</td><td>${co}</td></tr>`}).join("");}
function selectProject(p,row){
  selectedProject=p;
  if(selectedProjects.includes(p)){selectedProjects=selectedProjects.filter(x=>x!==p);row.classList.remove('selected');}
  else{selectedProjects.push(p);row.classList.add('selected');}
  if(selectedProjects.length>2){const first=selectedProjects.shift();document.querySelectorAll('#projectsTable tbody tr').forEach(r=>{if(r.cells[0]?.innerText===first)r.classList.remove('selected')});}
}
async function loadSelectedProject(){if(!selectedProject)return toast("Select project");const r=await fetch("/api/load-project/"+encodeURIComponent(selectedProject),{headers:authHeaders()});const d=await r.json();applyProject(d);calculate();showTab("steps");}
async function copySelectedProject(){if(!selectedProject)return toast("Select project");const r=await fetch("/api/load-project/"+encodeURIComponent(selectedProject),{headers:authHeaders()});const d=await r.json();d.project_name=(d.project_name||"Project")+" copy";const sr=await fetch("/api/save-project",{method:"POST",headers:apiHeaders(),body:JSON.stringify(d)});if(!sr.ok){const e=await sr.json();toast(e.detail||e.error||"Copy save error");return;}toast("Copied");refreshProjects();}
async function getProjectByFile(file){const r=await fetch('/api/load-project/'+encodeURIComponent(file),{headers:authHeaders()});return await r.json();}
async function _comparisonProjects(){
  if(selectedProjects.length>=2){return [await getProjectByFile(selectedProjects[0]), await getProjectByFile(selectedProjects[1])];}
  if(selectedProjects.length===1){return [collectProject(), await getProjectByFile(selectedProjects[0])];}
  toast('Karşılaştırma için Projects tablosundan 1 kayıt seçin; mevcut proje ile karşılaştırılır. İki kayıt seçerseniz iki kayıt karşılaştırılır.');
  return null;
}


async function compareSelectedProjects(){
  const projects=await _comparisonProjects(); if(!projects)return;
  const r=await fetch('/api/compare',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({projects})});
  const d=await r.json();
  $('comparisonBox').classList.remove('hidden');

  const rows=d.rows||[];
  const p1=d.project1||'Project 1';
  const p2=d.project2||'Project 2';
  const cards=rows.map(row=>{
    const cls=row.status==='equal'?'equal':(row.status==='better_p1'||row.status==='better_p2'?'good':'');
    const adv=row.advantage==='Equal'?'Equal':('Better: '+(row.advantage||'').replace('Project ','P'));
    return `
    <div class="v38CompareRow">
      <div class="v38CompareCard"><small>${row.metric}</small><b>${row.project1}</b><span>${p1}</span></div>
      <div class="v38CompareDelta ${cls}"><b>${row.diff||0}%</b><span>${adv}</span></div>
      <div class="v38CompareCard"><small>${row.metric}</small><b>${row.project2}</b><span>${p2}</span></div>
    </div>`}).join('');

  $('comparisonResult').innerHTML=`
    <div class="v38CompareDashboard">
      <div class="v38CompareTop">
        <div>
          <h2>Executive Comparison</h2>
          <p>${p1} vs ${p2}</p>
        </div>
        <button class="orange" onclick="exportComparisonPpt()">Export Compare PowerPoint</button>
      </div>
      <div class="v38CompareHead">
        <div>${p1}</div><div>Difference / Advantage</div><div>${p2}</div>
      </div>
      <div class="v38CompareRows">${cards}</div>
    </div>`;
}

async function exportComparisonPpt(){
  const projects=await _comparisonProjects(); if(!projects)return;
  try{
    toast('Grafikler program görünümünden hazırlanıyor...');
    const chart_pngs=[];
    for(const p of projects.slice(0,2)) chart_pngs.push(await projectChartPng(p));
    await postDownloadWithPayload('/api/export/comparison-ppt','DyeFlow_RS_Compare_Report.pptx',{projects,chart_pngs});
  }catch(e){
    toast('Comparison PPT export error: '+e.message);
  }
}


function selectFolderInfo(){toast("Web version saves inside backend/projects folder.");}
function loadLocalArchives(){try{machineArchive=JSON.parse(localStorage.getItem("dyeflow_machine_archive")||"[]");}catch(e){machineArchive=[];}try{chemicalArchive=JSON.parse(localStorage.getItem("dyeflow_chemical_archive")||"[]");}catch(e){chemicalArchive=[];}if(!machineArchive.length)machineArchive=[{machine_name:"600A",capacity_kg:600,total_power_used:30.60,drain_time_min:5}];}
function saveLocalArchives(){localStorage.setItem("dyeflow_machine_archive",JSON.stringify(machineArchive));localStorage.setItem("dyeflow_chemical_archive",JSON.stringify(chemicalArchive));}
function addCurrentMachineToArchive(){calculatePower();machineArchive.push({machine_name:val("machine_name"),capacity_kg:num("capacity_kg"),total_power_used:num("total_power_used"),drain_time_min:num("drain_time_min")});saveLocalArchives();renderArchives();toast("Machine added");}
function addProjectChemicalsToArchive(){steps.forEach(s=>s.chemicals.forEach(c=>chemicalArchive.push({supplier:c.supplier||"",chemical:c.chemical||"",amount:c.amount||0,unit:c.unit||"g/l",price:c.price||0})));saveLocalArchives();renderArchives();toast("Chemicals added");}
function renderArchives(){$("machineArchiveTable").querySelector("tbody").innerHTML=machineArchive.map((m,i)=>`<tr onclick="selectedMachineIndex=${i};highlightArchiveRow(this,'#machineArchiveTable')"><td>${m.machine_name||""}</td><td>${m.capacity_kg||0}</td><td>${m.total_power_used||0}</td><td>${m.drain_time_min||0}</td></tr>`).join("");$("chemicalArchiveTable").querySelector("tbody").innerHTML=chemicalArchive.map((c,i)=>`<tr onclick="selectedChemicalIndex=${i};highlightArchiveRow(this,'#chemicalArchiveTable')"><td>${c.supplier||""}</td><td>${c.chemical||""}</td><td>${c.amount||0}</td><td>${c.unit||""}</td><td>${c.price||0}</td></tr>`).join("");}
function highlightArchiveRow(row,table){document.querySelectorAll(table+" tbody tr").forEach(r=>r.classList.remove("selected"));row.classList.add("selected");}
function applySelectedMachine(){if(selectedMachineIndex===null)return toast("Select machine");const m=machineArchive[selectedMachineIndex];$("machine_name").value=m.machine_name||"";$("capacity_kg").value=m.capacity_kg||0;$("drain_time_min").value=m.drain_time_min||0;calculatePower();toast("Applied");}
function deleteSelectedMachine(){if(selectedMachineIndex===null)return toast("Select machine");machineArchive.splice(selectedMachineIndex,1);selectedMachineIndex=null;saveLocalArchives();renderArchives();}
function addSelectedChemicalToCurrentStep(){if(selectedChemicalIndex===null)return toast("Select chemical");if(!steps.length)addStep();const c=chemicalArchive[selectedChemicalIndex];steps[steps.length-1].chemicals.push({supplier:c.supplier||"",chemical:c.chemical||"",company:"",process:"",begin_c:steps[steps.length-1].beginning_temp,final_c:steps[steps.length-1].beginning_temp,dose_min:0,dose_time:0,circulation_time:0,amount:c.amount||0,unit:c.unit||"g/l",price:c.price||0});renderSteps();showTab("steps");toast("Added to current step");}
function deleteSelectedChemical(){if(selectedChemicalIndex===null)return toast("Select chemical");chemicalArchive.splice(selectedChemicalIndex,1);selectedChemicalIndex=null;saveLocalArchives();renderArchives();}
function selectReportImage(){$("reportImageInput").click();}
function handleReportImage(e){const file=e.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=()=>{$("reportImagePreview").innerHTML=`<img src="${reader.result}" style="max-width:320px;max-height:180px;border:1px solid #aaa;margin:8px 0;">`;};reader.readAsDataURL(file);}

function serializeCurrentChartSvg(){
  const svg=document.querySelector('#chart svg');
  if(!svg) return null;
  const clone=svg.cloneNode(true);
  clone.setAttribute('xmlns','http://www.w3.org/2000/svg');
  clone.setAttribute('width','1280');
  clone.setAttribute('height','470');
  return new XMLSerializer().serializeToString(clone);
}
function currentChartPngDataUrl(){
  return new Promise((resolve)=>{
    const svgText=serializeCurrentChartSvg();
    if(!svgText) return resolve(null);
    const svgBlob=new Blob([svgText],{type:'image/svg+xml;charset=utf-8'});
    const url=URL.createObjectURL(svgBlob);
    const img=new Image();
    img.onload=()=>{
      try{
        const canvas=document.createElement('canvas');
        canvas.width=1920; canvas.height=705;
        const ctx=canvas.getContext('2d');
        ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,canvas.width,canvas.height);
        ctx.drawImage(img,0,0,canvas.width,canvas.height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL('image/png'));
      }catch(e){URL.revokeObjectURL(url); resolve(null);}
    };
    img.onerror=()=>{URL.revokeObjectURL(url); resolve(null);};
    img.src=url;
  });
}
async function projectChartPng(project){
  const backup=JSON.parse(JSON.stringify(collectProject()));
  try{
    applyProject(JSON.parse(JSON.stringify(project)));
    await calculate();
    return await currentChartPngDataUrl();
  }finally{
    applyProject(backup);
    await calculate();
  }
}
async function postDownloadWithPayload(url,filename,payload){
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  if(!r.ok){let msg='Export error';try{const d=await r.json();msg=d.error||d.detail||msg;}catch(e){}toast(msg);return;}
  let b=await r.blob();let a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(a.href);
}

function postDownload(url,filename){fetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(collectProject())}).then(async r=>{if(!r.ok){toast("Export error");return;}let b=await r.blob();let a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=filename;a.click();URL.revokeObjectURL(a.href);});}
function downloadCSV(){calculate().then(()=>postDownload("/api/export/csv","DyeFlow_RS.csv"));}
function downloadPackage(){calculate().then(()=>postDownload("/api/export/package","DyeFlow_RS_Package.zip"));}
async function downloadPpt(){await calculate();const project=collectProject();const chart_png=await currentChartPngDataUrl();await postDownloadWithPayload('/api/export/powerpoint','DyeFlow_RS.pptx',{project,chart_png});}
function downloadReportHtml(){postDownload("/api/export/report-html","DyeFlow_RS_Report.html");}
async function downloadPDF(){
  await calculate();
  const project=collectProject();
  const chart_png=await currentChartPngDataUrl();
  fetch("/api/export/report-html",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({project,chart_png})})
    .then(r=>r.text()).then(html=>{const w=window.open("","_blank");w.document.open();w.document.write(html);w.document.close();setTimeout(()=>w.print(),500);});
}

async function downloadChartPNG(){
  await calculate();
  const project=collectProject();
  const chart_png=await currentChartPngDataUrl();
  await postDownloadWithPayload('/api/export/chart-png','DyeFlow_RS_Chart.png',{project,chart_png});
}


// ==============================
// v23 Local Project Manager
// Desktop-like local project save/load without login
// ==============================
const LOCAL_PROJECTS_KEY = "dyeflow_local_projects_v23";
const LOCAL_AUTOSAVE_KEY = "dyeflow_autosave_project_v23";

function sanitizeFilename(name){
  return String(name || "DyeFlow_Project").trim().replace(/[^A-Za-z0-9ğüşöçıİĞÜŞÖÇ _.-]+/g,"_").replace(/\s+/g,"_") || "DyeFlow_Project";
}

function nowStamp(){
  const d = new Date();
  const pad = n => String(n).padStart(2,"0");
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function projectPayload(){
  const p = collectProject();
  p._dyeflow_meta = {
    app: "DyeFlow RS",
    version: "v23_local_project_manager",
    saved_at: new Date().toISOString(),
    format: "dyeflow.project.json"
  };
  return p;
}

function downloadTextFile(filename, content, mime="application/json"){
  const blob = new Blob([content], {type: mime});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 250);
}

function getLocalProjects(){
  try { return JSON.parse(localStorage.getItem(LOCAL_PROJECTS_KEY) || "[]"); }
  catch(e){ return []; }
}

function setLocalProjects(list){
  localStorage.setItem(LOCAL_PROJECTS_KEY, JSON.stringify(list || []));
}

function upsertLocalProject(project){
  const list = getLocalProjects();
  const name = project.project_name || "Untitled Project";
  const key = sanitizeFilename(name).toLowerCase();
  const item = {
    key,
    project_name: name,
    company_name: project.company_name || "",
    updated: new Date().toLocaleString(),
    data: project
  };
  const idx = list.findIndex(x => x.key === key);
  if(idx >= 0) list[idx] = item; else list.unshift(item);
  setLocalProjects(list.slice(0, 50));
  renderLocalProjects();
}

function saveBrowserBackup(){
  const project = projectPayload();
  upsertLocalProject(project);
  localStorage.setItem(LOCAL_AUTOSAVE_KEY, JSON.stringify(project));
  const filename = `${sanitizeFilename(project.project_name || "DyeFlow_Project")}_${nowStamp()}.dyeflow.json`;
  downloadTextFile(filename, JSON.stringify(project, null, 2));
  toast("Project saved to browser and downloaded.");
}

function exportProjectJson(){
  const project = projectPayload();
  const filename = `${sanitizeFilename(project.project_name || "DyeFlow_Project")}_${nowStamp()}.json`;
  downloadTextFile(filename, JSON.stringify(project, null, 2));
  toast("JSON exported.");
}

function triggerImportProjectJson(){
  let inp = document.getElementById("projectJsonInput");
  if(!inp){
    inp = document.createElement("input");
    inp.type = "file";
    inp.id = "projectJsonInput";
    inp.accept = ".json,.dyeflow,.dyeflow.json,application/json";
    inp.style.display = "none";
    document.body.appendChild(inp);
    inp.addEventListener("change", handleImportProjectJson);
  }
  inp.value = "";
  inp.click();
}

function loadBrowserBackup(){
  triggerImportProjectJson();
}

function handleImportProjectJson(ev){
  const file = ev.target.files && ev.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try{
      const data = JSON.parse(reader.result);
      applyProject(data);
      upsertLocalProject(data);
      localStorage.setItem(LOCAL_AUTOSAVE_KEY, JSON.stringify(data));
      calculate();
      toast("Project loaded from file.");
      showTab("steps");
    }catch(e){
      console.error(e);
      toast("Invalid JSON project file.");
    }
  };
  reader.readAsText(file);
}

function renderLocalProjects(){
  const tbody = document.querySelector("#localProjectsTable tbody");
  if(!tbody) return;
  const list = getLocalProjects();
  if(!list.length){
    tbody.innerHTML = `<tr><td colspan="4">No local browser projects yet.</td></tr>`;
    return;
  }
  tbody.innerHTML = list.map((item, idx)=>`<tr onclick="loadLocalProjectByIndex(${idx})"><td>${item.project_name||"-"}</td><td>${item.company_name||"-"}</td><td>${item.updated||"-"}</td><td><button onclick="event.stopPropagation();deleteLocalProject(${idx})">Delete</button></td></tr>`).join("");
}

function loadLocalProjectByIndex(idx){
  const item = getLocalProjects()[idx];
  if(!item || !item.data) return toast("Local project not found.");
  applyProject(item.data);
  calculate();
  toast("Local project opened: " + (item.project_name || "Project"));
  showTab("steps");
}

function deleteLocalProject(idx){
  const list = getLocalProjects();
  list.splice(idx,1);
  setLocalProjects(list);
  renderLocalProjects();
  toast("Local project deleted.");
}

function autoSaveLocal(){
  try{
    const project = projectPayload();
    localStorage.setItem(LOCAL_AUTOSAVE_KEY, JSON.stringify(project));
  }catch(e){}
}

function loadAutoSave(){
  try{
    const raw = localStorage.getItem(LOCAL_AUTOSAVE_KEY);
    if(!raw) return toast("No autosave found.");
    const data = JSON.parse(raw);
    applyProject(data);
    calculate();
    toast("Autosave loaded.");
  }catch(e){ toast("Autosave could not be loaded."); }
}

async function checkHealth(){
  try{
    const r = await fetch("/api/health");
    const d = await r.json();
    toast("Health: " + (d.status || "ok"));
  }catch(e){
    toast("Health check failed.");
  }
}

// Start local autosave after UI is ready
window.addEventListener("load",()=>{
  renderLocalProjects();
  setInterval(autoSaveLocal, 30000);
});
