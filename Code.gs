const CONFIG = {
  DATA_FOLDER_ID: '1QLm1cHNVKQSx5vpu7Rx1PrjjyEYG1eUM',
  MASTER_INDEX_FILE: 'Master_Category_Index.json',
  CACHE_SECONDS: 300
};

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // CORS Headers are automatically handled by Google Apps Script ContentService
  try {
    let action = e.parameter.action;
    let data;

    if (e.postData && e.postData.contents) {
      try {
        let body = JSON.parse(e.postData.contents);
        if (body.action) action = body.action;
        data = body;
      } catch (err) {}
    }

    if (!action) {
      return ContentService.createTextOutput(JSON.stringify({ error: "No action provided" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let result;
    if (action === 'getCategoriesList') {
      result = getCategoriesList();
    } else if (action === 'getCategoryProducts') {
      result = getCategoryProducts(e.parameter.categoryName || (data && data.categoryName));
    } else if (action === 'requestL1Update') {
      result = requestL1Update(data.url, data.category);
    } else if (action === 'exportProductsJson') {
      result = exportProductsJson(data);
    } else if (action === 'getCategoryCountsBatch') {
      result = getCategoryCountsBatch(data.names);
    } else if (action === 'forceRefreshCache') {
      result = forceRefreshCache();
    } else {
      return ContentService.createTextOutput(JSON.stringify({ error: "Unknown action" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message, success: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getCategoriesList() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
    const byKey = {};
    const out = [];

    function addCategory(cat) {
      cat = normalizeCategoryIndexItem_(cat);
      if (!cat.name) return;

      const key = String(cat.fileName || cat.name).toLowerCase();
      if (byKey[key]) {
        // Merge stronger metadata from Master Index with Drive file discovery.
        if (!byKey[key].productCount && cat.productCount) byKey[key].productCount = cat.productCount;
        if ((!byKey[key].updated || byKey[key].updated === '09.05.2026 · 17:59') && cat.updated) byKey[key].updated = cat.updated;
        if (!byKey[key].fileName && cat.fileName) byKey[key].fileName = cat.fileName;
        return;
      }

      byKey[key] = cat;
      out.push(cat);
    }

    // 1) Read Master_Category_Index.json when available.
    // 2) Also scan the folder and merge all JSON files so incomplete master index never limits the UI to 5/50 items.
    try {
      const index = getMasterCategoryIndex_();
      if (Array.isArray(index)) {
        index.forEach(function(item) { addCategory(item); });
      }
    } catch (idxErr) {}

    const files = folder.searchFiles("title contains '.json' and trashed = false");
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (!fileName.toLowerCase().endsWith('.json')) continue;
      if (fileName.toLowerCase() === CONFIG.MASTER_INDEX_FILE.toLowerCase()) continue;

      addCategory({
        name: fileName.replace(/\.json$/i, '').replace(/_/g, ' '),
        fileName: fileName,
        productCount: 0,
        updated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd.MM.yyyy · HH:mm')
      });
    }

    out.sort(function(a, b) {
      return String(a.name).localeCompare(String(b.name), undefined, { numeric: true, sensitivity: 'base' });
    });

    return out;
  } catch (e) {
    return { error: e.toString() };
  }
}

function normalizeCategoryIndexItem_(item) {
  if (!item || typeof item !== 'object') item = { name: String(item || '') };

  const rawFile = String(
    item.fileName || item.filename || item.file_name || item.file || item.jsonFile || item.json_file || item.path || ''
  ).trim();

  let rawName = String(
    item.name || item.category || item.categoryName || item.category_name || item.title || item.displayName || item.display_name || ''
  ).trim();

  if (!rawName && rawFile) rawName = rawFile.replace(/\.json$/i, '');

  rawName = rawName
    .replace(/\.json$/i, '')
    .replace(/_/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const fileName = rawFile
    ? rawFile.split('/').pop()
    : (rawName ? rawName.replace(/\s+/g, '_') + '.json' : '');

  const productCount = Number(
    item.productCount || item.product_count || item.products || item.productTotal || item.product_total || item.count || item.totalProducts || item.total_products || item.total || 0
  );

  return {
    name: rawName,
    fileName: fileName,
    productCount: isNaN(productCount) ? 0 : productCount,
    updated: String(item.updated || item.lastUpdated || item.last_updated || item.modified || item.modifiedAt || item.modified_at || '09.05.2026 · 17:59')
  };
}

function buildMasterIndexMap_() {
  const out = {};

  try {
    const file = findFileByName_(CONFIG.MASTER_INDEX_FILE);
    if (!file) return out;

    const content = file.getBlob().getDataAsString();
    const data = JSON.parse(content);
    const index = Array.isArray(data) ? data : (data.categories || data.data || []);

    index.forEach(function(item) {
      const cat = normalizeCategoryIndexItem_(item);
      if (!cat.name) return;
      const fileBase = String(cat.fileName || '').replace(/\.json$/i, '').trim();
      const clean = cat.name.replace(/_/g, ' ').trim();
      out[cat.name.toLowerCase()] = cat;
      out[clean.toLowerCase()] = cat;
      if (cat.fileName) out[cat.fileName.toLowerCase()] = cat;
      if (fileBase) out[fileBase.toLowerCase()] = cat;
      if (fileBase) out[fileBase.replace(/_/g, ' ').toLowerCase()] = cat;
    });
  } catch (e) {}

  return out;
}

function getCategoryProducts(categoryName) {
  if (!categoryName) {
    return { error: 'Category name missing.' };
  }

  try {
    const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
    let matchedFile = findCategoryFileForExport_(folder, categoryName);

    if (!matchedFile) {
      const safeCatName = String(categoryName).replace(/'/g, "\\'");
      const files = folder.searchFiles("title contains '" + safeCatName + "' and trashed = false");

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        if (name.toLowerCase().endsWith('.json')) {
          const cleanName = name.replace(/\.json$/i, '').replace(/_/g, ' ');
          if (cleanName.toLowerCase() === String(categoryName).toLowerCase()) {
            matchedFile = file;
            break;
          }
        }
      }
    }

    if (!matchedFile) {
      return { error: 'Category JSON file not found for: ' + categoryName };
    }

    const content = matchedFile.getBlob().getDataAsString();
    const data = JSON.parse(content);

    let products = [];
    if (Array.isArray(data)) {
      products = data;
    } else if (data && typeof data === 'object') {
      if (Array.isArray(data.products)) products = data.products;
      else if (Array.isArray(data.items)) products = data.items;
      else if (Array.isArray(data.data)) products = data.data;
      else {
        let largestArray = [];
        function findArrays(obj) {
          if (!obj || typeof obj !== 'object') return;
          if (Array.isArray(obj)) {
            if (obj.length > largestArray.length && obj.length > 0 && typeof obj[0] === 'object') {
              largestArray = obj;
            }
          } else {
            for (let k in obj) {
              findArrays(obj[k]);
            }
          }
        }
        findArrays(data);
        products = largestArray;
      }
    }

    return products;
  } catch (e) {
     return { error: e.toString() };
  }
}

function searchCategories(query) {
  query = String(query || '').toLowerCase().trim();

  const index = getMasterCategoryIndex_();

  if (!query) {
    return index.slice(0, 50);
  }

  return index
    .filter(item => {
      const text = [
        item.category,
        item.name,
        item.title,
        item.fileName,
        item.keywords
      ].join(' ').toLowerCase();

      return text.includes(query);
    })
    .slice(0, 50);
}

function getProductsByCategory(categoryFileName) {
  if (!categoryFileName) {
    throw new Error('Category file name missing.');
  }

  const safeFileName = String(categoryFileName).replace(/[\/\\]/g, '');
  const cacheKey = 'products_' + safeFileName;

  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const file = findFileByName_(safeFileName);
  if (!file) {
    throw new Error('Category JSON file not found: ' + safeFileName);
  }

  const content = file.getBlob().getDataAsString();
  const data = JSON.parse(content);

  const products = Array.isArray(data)
    ? data
    : data.products || data.items || data.data || [];

  cache.put(cacheKey, JSON.stringify(products), CONFIG.CACHE_SECONDS);

  return products;
}

function getMasterCategoryIndex_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('master_category_index');

  if (cached) {
    return JSON.parse(cached);
  }

  const file = findFileByName_(CONFIG.MASTER_INDEX_FILE);
  if (!file) {
    throw new Error('Master_Category_Index.json not found.');
  }

  const content = file.getBlob().getDataAsString();
  const data = JSON.parse(content);

  const index = Array.isArray(data)
    ? data
    : data.categories || data.data || [];

  cache.put('master_category_index', JSON.stringify(index), CONFIG.CACHE_SECONDS);

  return index;
}

function findFileByName_(fileName) {
  const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
  const files = folder.getFilesByName(fileName);

  if (files.hasNext()) {
    return files.next();
  }

  return null;
}

function forceRefreshCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('master_category_index');
  // Return success
  return true;
}

/**
 * Fast brand/category JSON export for the portal.
 * IMPORTANT FIX:
 * It does NOT scan all 9,000+ JSON files anymore.
 * It opens only the exact category file names sent by the frontend.
 */
function exportProductsJson(payload) {
  payload = payload || {};

  const categoryName = String(payload.categoryName || '').trim();
  const categoryQuery = String(payload.categoryQuery || '').trim();
  const brandQuery = String(payload.brandQuery || '').trim();
  const visibleCategoryNames = Array.isArray(payload.categoryNames) ? payload.categoryNames : [];

  let targetCategoryNames = [];

  if (categoryName) {
    targetCategoryNames = [categoryName];
  } else if (visibleCategoryNames.length) {
    targetCategoryNames = visibleCategoryNames;
  } else {
    return {
      success: false,
      error: 'Please search or select a category first. Export stopped to avoid scanning all JSON files.'
    };
  }

  targetCategoryNames = uniqueStringsForExport_(targetCategoryNames)
    .filter(function(name) { return !!String(name || '').trim(); });

  if (!targetCategoryNames.length) {
    return { success: false, error: 'No matching category found for export.' };
  }

  if (targetCategoryNames.length > 50) {
    return {
      success: false,
      error: 'Too many category files selected for one export (' + targetCategoryNames.length + '). Please refine the category search, for example Currency Counting instead of Currency.'
    };
  }

  try {
    const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
    const products = [];
    const seenProducts = {};
    const matchedFiles = [];
    const missingFiles = [];

    let openedFiles = 0;
    let matchedCategoryFiles = 0;

    targetCategoryNames.forEach(function(catName) {
      const file = findCategoryFileForExport_(folder, catName);

      if (!file) {
        missingFiles.push(catName);
        return;
      }

      openedFiles++;
      matchedCategoryFiles++;

      let data;
      try {
        data = JSON.parse(file.getBlob().getDataAsString());
      } catch (err) {
        missingFiles.push(catName + ' (invalid JSON)');
        return;
      }

      const fileProducts = extractProductArrayForExport_(data);
      let fileMatchCount = 0;
      const normalizedBrand = normalizeExportText_(brandQuery);

      fileProducts.forEach(function(product) {
        if (!product || typeof product !== 'object') return;
        if (normalizedBrand && !productMatchesBrandForExport_(product, normalizedBrand)) return;

        const uniqueKey = String(
          product.product_id ||
          product.product_url ||
          product.url ||
          product.link ||
          JSON.stringify(product).slice(0, 500)
        );

        if (seenProducts[uniqueKey]) return;
        seenProducts[uniqueKey] = true;

        products.push(product);
        fileMatchCount++;
      });

      if (fileMatchCount) {
        matchedFiles.push({ fileName: file.getName(), count: fileMatchCount });
      }
    });

    return {
      success: true,
      products: products,
      total: products.length,
      brandQuery: brandQuery,
      categoryQuery: categoryName || categoryQuery || 'Category_Search',
      scannedFiles: openedFiles,
      matchedCategoryFiles: matchedCategoryFiles,
      missingFiles: missingFiles,
      matchedFiles: matchedFiles,
      fileName: buildExportFileName_(categoryName || categoryQuery || 'Category_Search', brandQuery)
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function findCategoryFileForExport_(folder, categoryName) {
  const candidates = buildCategoryFileNameCandidates_(categoryName);

  for (let i = 0; i < candidates.length; i++) {
    const files = folder.getFilesByName(candidates[i]);
    if (files.hasNext()) return files.next();
  }

  return null;
}

function buildCategoryFileNameCandidates_(categoryName) {
  const raw = String(categoryName || '').trim();
  const withoutJson = raw.replace(/\.json$/i, '');
  const underscore = withoutJson.replace(/\s+/g, '_');
  const space = withoutJson.replace(/_/g, ' ');

  return uniqueStringsForExport_([
    raw,
    withoutJson + '.json',
    underscore + '.json',
    space.replace(/\s+/g, '_') + '.json',
    sanitizeExportFilePart_(withoutJson) + '.json'
  ]);
}

function uniqueStringsForExport_(arr) {
  const seen = {};
  const out = [];

  (arr || []).forEach(function(item) {
    const value = String(item || '').trim();
    if (!value) return;
    const key = value.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push(value);
  });

  return out;
}

function productMatchesBrandForExport_(product, normalizedBrand) {
  if (!normalizedBrand) return true;

  const productBrand = normalizeExportText_(getExportProductBrand_(product));
  const productTitle = normalizeExportText_(product.title || product.product_name || product.name || product.productName || '');
  const productMake = normalizeExportText_(product.make || product.manufacturer || product.brand_name || '');
  const productJson = normalizeExportText_(JSON.stringify(product));

  const haystack = normalizeExportText_([
    productBrand,
    stripGemRegisteredSuffix_(productBrand),
    productTitle,
    productMake,
    productJson
  ].join(' '));

  const query = normalizedBrand;
  const looseQuery = normalizeForLooseMatch_(query);
  const looseQueryWithoutR = stripGemRegisteredSuffix_(looseQuery);
  const looseHaystack = normalizeForLooseMatch_(haystack);

  if (haystack.indexOf(query) !== -1) return true;
  if (looseHaystack.indexOf(looseQuery) !== -1) return true;
  if (looseQueryWithoutR && looseHaystack.indexOf(looseQueryWithoutR) !== -1) return true;

  const productBrandLoose = normalizeForLooseMatch_(productBrand);
  const productBrandWithoutR = stripGemRegisteredSuffix_(productBrandLoose);

  if (productBrandLoose.indexOf(looseQuery) !== -1) return true;
  if (productBrandWithoutR.indexOf(looseQuery) !== -1) return true;
  if (looseQueryWithoutR && productBrandLoose.indexOf(looseQueryWithoutR) !== -1) return true;

  const tokens = query.split(' ').filter(function(token) { return token.length >= 2; });
  if (tokens.length && tokens.every(function(token) { return haystack.indexOf(token) !== -1; })) return true;

  return false;
}

function extractProductArrayForExport_(data) {
  if (Array.isArray(data)) return data;

  if (data && typeof data === 'object') {
    if (Array.isArray(data.products)) return data.products;
    if (Array.isArray(data.items)) return data.items;
    if (Array.isArray(data.data)) return data.data;

    let largestArray = [];

    function findArrays(obj) {
      if (!obj || typeof obj !== 'object') return;

      if (Array.isArray(obj)) {
        if (obj.length > largestArray.length && obj.length > 0 && typeof obj[0] === 'object') {
          largestArray = obj;
        }
        return;
      }

      Object.keys(obj).forEach(function(key) {
        findArrays(obj[key]);
      });
    }

    findArrays(data);
    return largestArray;
  }

  return [];
}

function getExportProductBrand_(product) {
  return String(
    product.brand ||
    product.brand_name ||
    product.manufacturer ||
    product.make ||
    product.oem ||
    product.oem_name ||
    ''
  ).trim();
}

function normalizeExportText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[_\-]+/g, ' ')
    .replace(/®|™/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeForLooseMatch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/®|™/g, '')
    .replace(/[^a-z0-9]+/g, '')
    .trim();
}

function stripGemRegisteredSuffix_(value) {
  let text = String(value || '').trim();
  while (text.length > 3 && text.endsWith('r')) {
    text = text.slice(0, -1);
  }
  return text;
}

function buildExportFileName_(category, brand) {
  const c = sanitizeExportFilePart_(category || 'Category');
  const b = brand ? '_' + sanitizeExportFilePart_(brand) : '_Full';
  return c + b + '_Filtered.json';
}

function sanitizeExportFilePart_(value) {
  return String(value || '')
    .replace(/\.json$/i, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80) || 'Export';
}

/**
 * Counts products only for the currently visible category cards.
 * This keeps the UI accurate without scanning/parsing every JSON file on first load.
 */
function getCategoryCountsBatch(items) {
  items = Array.isArray(items) ? items : [];
  const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
  const cache = CacheService.getScriptCache();
  const counts = {};

  items.slice(0, 60).forEach(function(item) {
    const name = String((item && item.name) || '').trim();
    const fileName = String((item && item.fileName) || '').trim();
    if (!name && !fileName) return;

    const cacheKey = 'cat_count_' + Utilities.base64EncodeWebSafe(fileName || name).slice(0, 80);
    const cached = cache.get(cacheKey);
    if (cached !== null) {
      counts[name || fileName] = Number(cached || 0);
      return;
    }

    let count = 0;
    try {
      const file = fileName ? findCategoryFileForExport_(folder, fileName) : findCategoryFileForExport_(folder, name);
      if (file) {
        const data = JSON.parse(file.getBlob().getDataAsString());
        count = extractProductArrayForExport_(data).length;
      }
    } catch (e) {
      count = 0;
    }

    cache.put(cacheKey, String(count), CONFIG.CACHE_SECONDS);
    counts[name || fileName] = count;
  });

  return { success: true, counts: counts };
}

function requestL1Update(productUrl, categoryName) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DATA_FOLDER_ID);
    let tasksFolder;
    const folders = folder.getFoldersByName("Tasks");
    if (folders.hasNext()) {
      tasksFolder = folders.next();
    } else {
      tasksFolder = folder.createFolder("Tasks");
    }
    
    const fileName = 'task_L1_Update_' + new Date().getTime() + '.json';
    const taskData = {
      productUrl: productUrl,
      categoryName: categoryName,
      timestamp: new Date().getTime(),
      status: 'pending'
    };
    
    tasksFolder.createFile(fileName, JSON.stringify(taskData), MimeType.PLAIN_TEXT);
    return { success: true, message: 'L1 Realtime Sync Triggered! The background engine is now fetching live data.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
