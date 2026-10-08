import re, os, sys

sys.stdout.reconfigure(encoding='utf-8')

html_path = '591看屋物件地圖與比較分析.html'

with open(html_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Insert URL Auto-Import Box at start of <main>
url_box_html = """  <!-- Main Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-7">

    <!-- 網址貼上與自動入庫專區 (URL Auto-Import to PostgreSQL) -->
    <section class="bg-gradient-to-r from-blue-900 via-indigo-950 to-slate-900 rounded-2xl shadow-xl p-5 md:p-6 text-white border border-indigo-500/40">
      <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div class="space-y-1">
          <div class="flex items-center gap-2">
            <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-400 text-slate-950 flex items-center gap-1">
              <i class="fa-solid fa-wand-magic-sparkles"></i> 自動擷取入庫
            </span>
            <h3 class="text-lg md:text-xl font-extrabold text-white">貼上房仲網址 ➜ 自動擷取並加入 PostgreSQL 資料庫</h3>
          </div>
          <p class="text-xs text-indigo-200">
            支援 <strong>591 售屋網、信義房屋、永義房屋、有巢氏房屋、永慶房屋</strong> 等網站。系統自動分析格局、價格、坪數，並下載現場照片存入本地資料夾，即時加入「待看物件」！
          </p>
        </div>
        <div class="flex-1 max-w-2xl flex flex-col sm:flex-row gap-2">
          <input type="url" id="inputPropertyUrl" placeholder="請在此貼上房屋網頁網址 (例: https://sale.591.com.tw/... 或 https://www.sinyi.com.tw/...)" 
                 class="flex-1 px-4 py-2.5 rounded-xl bg-slate-800/90 text-white placeholder-slate-400 border border-indigo-400/50 text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-amber-400">
          <button onclick="submitPropertyUrl()" id="btnSubmitUrl" 
                  class="px-5 py-2.5 bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 font-extrabold rounded-xl text-xs sm:text-sm flex items-center justify-center gap-2 shadow-md hover:shadow-lg transition shrink-0">
            <i class="fa-solid fa-cloud-arrow-down"></i> 擷取並加入待看
          </button>
        </div>
      </div>
      <div id="crawlStatusMsg" class="hidden mt-3 text-xs p-3 rounded-xl bg-indigo-950/90 border border-indigo-400/50 text-indigo-200 flex items-center gap-3">
        <i class="fa-solid fa-spinner fa-spin text-amber-400 text-base"></i>
        <span id="crawlStatusText">正在連接網站擷取資料並下載照片至本地...</span>
      </div>
    </section>"""

if '<!-- 網址貼上與自動入庫專區' not in content:
    content = content.replace('<!-- Main Container -->\n  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-7">', url_box_html)

# 2. Add Delete button in popup actions
popup_buttons_old = """            <!-- Quick Move in Popup -->
            <div style="display:flex; gap:4px; margin-bottom:8px;">
              <button onclick="movePropertyCategory(${p.id}, 'consider')" style="flex:1; padding:4px; background:#10b981; color:white; border:none; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;">❤️ 列入考慮</button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" style="flex:1; padding:4px; background:#ef4444; color:white; border:none; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;">✕ 不考慮</button>
            </div>"""

popup_buttons_new = """            <!-- Quick Move in Popup -->
            <div style="display:flex; gap:4px; margin-bottom:8px;">
              <button onclick="movePropertyCategory(${p.id}, 'consider')" style="flex:1; padding:4px; background:#10b981; color:white; border:none; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;" title="列入考慮">❤️ 考慮</button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" style="flex:1; padding:4px; background:#64748b; color:white; border:none; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;" title="不考慮">✕ 排除</button>
              <button onclick="deleteProperty(${p.id}, '${p.name}')" style="padding:4px 8px; background:#ef4444; color:white; border:none; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;" title="從資料庫刪除"><i class="fa-solid fa-trash-can"></i></button>
            </div>"""

content = content.replace(popup_buttons_old, popup_buttons_new)

# 3. Add Delete button in table actions
old_action_pending = """            <div class="flex items-center justify-center gap-1.5">
              <button onclick="movePropertyCategory(${p.id}, 'consider')" class="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold text-[11px] flex items-center gap-1 shadow-sm transition" title="看完列入考慮 (同步搬移檔案)">
                <i class="fa-solid fa-heart"></i> 列入考慮
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" class="px-2 py-1 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg font-bold text-[11px] flex items-center gap-1 transition" title="看完不考慮 (同步搬移檔案)">
                <i class="fa-solid fa-xmark"></i> 不考慮
              </button>
            </div>"""

new_action_pending = """            <div class="flex items-center justify-center gap-1.5">
              <button onclick="movePropertyCategory(${p.id}, 'consider')" class="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold text-[11px] flex items-center gap-1 shadow-sm transition" title="看完列入考慮 (寫入資料庫)">
                <i class="fa-solid fa-heart"></i> 列入考慮
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" class="px-2 py-1 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg font-bold text-[11px] flex items-center gap-1 transition" title="看完不考慮 (寫入資料庫)">
                <i class="fa-solid fa-xmark"></i> 不考慮
              </button>
              <button onclick="deleteProperty(${p.id}, '${p.name}')" class="px-2 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg font-bold text-[11px] flex items-center gap-1 transition" title="從資料庫徹底刪除">
                <i class="fa-solid fa-trash-can"></i> 刪除
              </button>
            </div>"""

content = content.replace(old_action_pending, new_action_pending)

old_action_consider = """            <div class="flex items-center justify-center gap-1.5">
              <span class="badge bg-emerald-100 text-emerald-800 font-bold mr-1"><i class="fa-solid fa-heart text-emerald-600 mr-1"></i>已列考慮</span>
              <button onclick="movePropertyCategory(${p.id}, 'pending')" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-[11px] transition" title="移回待看">
                <i class="fa-solid fa-rotate-left"></i> 回待看
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded text-[11px] transition" title="改為不考慮">
                <i class="fa-solid fa-xmark"></i> 不考慮
              </button>
            </div>"""

new_action_consider = """            <div class="flex items-center justify-center gap-1.5">
              <span class="badge bg-emerald-100 text-emerald-800 font-bold mr-1"><i class="fa-solid fa-heart text-emerald-600 mr-1"></i>已列考慮</span>
              <button onclick="movePropertyCategory(${p.id}, 'pending')" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-[11px] transition" title="移回待看">
                <i class="fa-solid fa-rotate-left"></i> 回待看
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'rejected')" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded text-[11px] transition" title="改為不考慮">
                <i class="fa-solid fa-xmark"></i> 不考慮
              </button>
              <button onclick="deleteProperty(${p.id}, '${p.name}')" class="px-2 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg font-bold text-[11px] flex items-center gap-1 transition" title="從資料庫徹底刪除">
                <i class="fa-solid fa-trash-can"></i> 刪除
              </button>
            </div>"""

content = content.replace(old_action_consider, new_action_consider)

old_action_rejected = """            <div class="flex items-center justify-center gap-1.5">
              <span class="badge bg-slate-200 text-slate-600 font-medium mr-1"><i class="fa-solid fa-ban mr-1"></i>不考慮</span>
              <button onclick="movePropertyCategory(${p.id}, 'pending')" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-[11px] transition" title="移回待看">
                <i class="fa-solid fa-rotate-left"></i> 回待看
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'consider')" class="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[11px] transition" title="改為考慮">
                <i class="fa-solid fa-heart"></i> 改考慮
              </button>
            </div>"""

new_action_rejected = """            <div class="flex items-center justify-center gap-1.5">
              <span class="badge bg-slate-200 text-slate-600 font-medium mr-1"><i class="fa-solid fa-ban mr-1"></i>不考慮</span>
              <button onclick="movePropertyCategory(${p.id}, 'pending')" class="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-[11px] transition" title="移回待看">
                <i class="fa-solid fa-rotate-left"></i> 回待看
              </button>
              <button onclick="movePropertyCategory(${p.id}, 'consider')" class="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded text-[11px] transition" title="改為考慮">
                <i class="fa-solid fa-heart"></i> 改考慮
              </button>
              <button onclick="deleteProperty(${p.id}, '${p.name}')" class="px-2 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg font-bold text-[11px] flex items-center gap-1 transition" title="從資料庫徹底刪除">
                <i class="fa-solid fa-trash-can"></i> 刪除
              </button>
            </div>"""

content = content.replace(old_action_rejected, new_action_rejected)

# 4. Insert JavaScript functions: deleteProperty, submitPropertyUrl, loadPropertiesFromDb
new_js = """
    // Delete Property from Database
    async function deleteProperty(id, name) {
      if (!confirm(`確定要將物件「${name}」從 PostgreSQL 資料庫徹底刪除嗎？\\n（此動作會將資料庫記錄與關聯圖片一併刪除）`)) {
        return;
      }
      try {
        const res = await fetch(`http://127.0.0.1:8899/api/delete?id=${id}`, { method: 'POST' });
        const data = await res.json();
        if (data.success) {
          showToast(`🗑️ <strong>已從資料庫刪除</strong>：已徹底移除「${name}」！`, '<i class="fa-solid fa-trash-can text-rose-400"></i>');
          allProperties = allProperties.filter(x => x.id !== id);
          delete propertyCategories[id];
          applyFiltersAndSort();
        } else {
          alert('刪除失敗: ' + (data.error || '未知錯誤'));
        }
      } catch (e) {
        console.error('Delete failed:', e);
        alert('無法連線至本機後端服務: ' + e);
      }
    }

    // Submit Property URL for Auto-Crawling and DB Insertion
    async function submitPropertyUrl() {
      const input = document.getElementById('inputPropertyUrl');
      const btn = document.getElementById('btnSubmitUrl');
      const statusBox = document.getElementById('crawlStatusMsg');
      const statusText = document.getElementById('crawlStatusText');
      const url = input.value.trim();

      if (!url) {
        alert('請先貼上房屋網址！');
        input.focus();
        return;
      }

      btn.disabled = true;
      btn.classList.add('opacity-50');
      statusBox.classList.remove('hidden');
      statusText.innerHTML = `正在從網址擷取資料，並自動下載現場照片至「物件圖片」資料夾中，請稍候...`;

      try {
        const res = await fetch('http://127.0.0.1:8899/api/crawl', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url })
        });
        const data = await res.json();
        if (data.success) {
          input.value = '';
          statusBox.classList.add('hidden');
          showToast(`🎉 <strong>【資料庫已建立】</strong><br>${data.message}`, '<i class="fa-solid fa-circle-check text-emerald-400"></i>');
          await loadPropertiesFromDb();
          switchCategoryTab('pending');
        } else {
          alert('擷取失敗: ' + (data.error || '伺服器未回應'));
          statusBox.classList.add('hidden');
        }
      } catch (e) {
        alert('連線失敗: 請確認本機後端服務正在運行 (http://127.0.0.1:8899/)');
        statusBox.classList.add('hidden');
      } finally {
        btn.disabled = false;
        btn.classList.remove('opacity-50');
      }
    }

    // Load Properties Dynamically from PostgreSQL taichung_house
    async function loadPropertiesFromDb() {
      try {
        const res = await fetch('http://127.0.0.1:8899/api/properties');
        if (res.ok) {
          const data = await res.json();
          if (data.success && data.properties && data.properties.length > 0) {
            allProperties = data.properties.map(p => ({
              id: p.id,
              code: p.code,
              name: p.community_name,
              title: p.title,
              price: p.price_total,
              unitPrice: p.unit_price,
              layout: p.layout_raw || `${p.rooms}房${p.living_rooms}廳${p.bathrooms}衛`,
              rooms: p.rooms,
              baths: p.bathrooms,
              floor: p.floor_info || `${p.current_floor}/${p.total_floors}F`,
              floorNum: p.current_floor || 5,
              totalFloors: p.total_floors || 10,
              age: (p.age ? p.age + '年' : '未載明'),
              ageNum: p.age || 30,
              totalArea: p.total_area,
              mainArea: p.indoor_area || 0,
              subArea: p.attached_area || 0,
              indoor: p.indoor_total || 25,
              parkingType: p.parking_type || '無車位',
              parkingNote: p.parking_desc || p.parking_type || '無車位',
              mgmtFee: p.management_fee || '未載明',
              mgmtFeeNum: parseInt((p.management_fee || '').replace(/[^0-9]/g, '')) || 0,
              publicRatio: (p.public_ratio ? p.public_ratio + '%' : '未載明'),
              publicRatioNum: p.public_ratio || 25,
              publicRatioDetail: p.public_ratio_desc || '',
              district: p.district || '南屯區',
              area: p.district || '台中市',
              orientation: p.orientation || '未載明',
              note: p.title || '',
              lat: p.lat,
              lng: p.lng,
              url: p.original_url || '',
              category: p.decision_status || 'pending',
              themeCategory: (p.parking_type && p.parking_type.includes('平面') ? 'plane' : 'view'),
              source: p.original_source_platform || '591',
              imgCount: p.img_count || 0,
              sampleImg: p.sample_img || ''
            }));

            propertyCategories = {};
            allProperties.forEach(p => {
              propertyCategories[p.id] = p.category || 'pending';
            });

            applyFiltersAndSort();
            const badge = document.getElementById('syncServiceBadge');
            if (badge) {
              badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-400/40';
              badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-400 inline-block"></span> PostgreSQL 16 資料庫連線中 (${allProperties.length} 戶)`;
            }
            return true;
          }
        }
      } catch (e) {
        console.warn('Could not load from DB API, using fallback embedded data', e);
      }
      return false;
    }
"""

# Replace movePropertyCategory to write directly to DB
old_move_code = """      // Try triggering real-time disk move via HouseSyncService
      if (isLocalServerAvailable) {
        try {
          const res = await fetch(`http://127.0.0.1:8899/api/move?id=${id}&to=${targetCategory}`);
          if (res.ok) {
            const data = await res.json();
            showToast(`🟢 <strong>【Google Drive 同步成功】</strong><br>已將「${p.name}」移至「${catNames[targetCategory]}」資料夾！`, '<i class="fa-solid fa-folder-check text-emerald-400"></i>');
            return;
          }
        } catch (e) {
          console.warn('Sync server call failed', e);
        }
      }"""

new_move_code = """      // Direct write to PostgreSQL taichung_house via backend API
      try {
        const res = await fetch(`http://127.0.0.1:8899/api/move?id=${id}&to=${targetCategory}`, { method: 'POST' });
        if (res.ok) {
          const data = await res.json();
          showToast(`🟢 <strong>【PostgreSQL 更新成功】</strong><br>已將「${p.name}」狀態即時寫入資料庫「${catNames[targetCategory]}」！`, '<i class="fa-solid fa-database text-emerald-400"></i>');
          return;
        }
      } catch (e) {
        console.warn('DB update failed', e);
      }"""

content = content.replace(old_move_code, new_move_code)

# Add loadPropertiesFromDb() in initialize
init_old = """    // Initialize
    applyFiltersAndSort();
    checkLocalSyncService();"""

init_new = """    // Initialize
    loadPropertiesFromDb().then(loaded => {
      if (!loaded) {
        applyFiltersAndSort();
      }
      checkLocalSyncService();
    });"""

content = content.replace(init_old, init_new)

# Insert the new functions before initialize
content = content.replace('    // Initialize', new_js + '\n    // Initialize')

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Successfully updated 591看屋物件地圖與比較分析.html!")
