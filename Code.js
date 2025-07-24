/**
 * Document Center - Main Backend Code (RESOLVED VERSION)
 * Google Apps Script backend for document management system
 */

// Configuration
const CONFIG = {
  appName: 'Document Center',
  version: '1.0.0',
  sessionTimeout: 30, // minutes
  animalAvatars: ['🐺', '🦊', '🐨', '🐸', '🦋', '🐧', '🦁', '🐯', '🐼', '🐰', '🦄', '🐙', '🦉', '🐢', '🦆', '🦅', '🦜', '🦩'],
  maxSearchResults: 100,
  recentActivityLimit: 20
};

// =====================================
// WEB APP ENTRY POINT
// =====================================

/**
 * Main entry point for the web application
 */
function doGet(e) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return createAuthRequiredPage();
    }
    trackUserSession(user, 'login');
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle(CONFIG.appName);
  } catch (error) {
    console.error('Error in doGet:', error);
    return createErrorPage(error.toString());
  }
}

/**
 * Include external HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =====================================
// INITIAL DATA LOADING
// =====================================

function getInitialData() {
  try {
    const user = getCurrentUser();
    if (!user || !user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    
    const response = {
      success: true,
      user: user,
      documents: getDocuments().documents || [],
      categories: getCategories().categories || [],
      tags: getTags().tags || [],
      favorites: getUserFavorites().favoriteIds || [],
      recentActivity: getRecentActivity().activities || [],
      onlineUsers: getOnlineUsers().users || [],
      analytics: getAnalyticsData().analytics || {},
      config: { appName: CONFIG.appName, version: CONFIG.version }
    };
    
    return JSON.stringify(response);
  } catch (error) {
    console.error('SERVER: CRITICAL ERROR in getInitialData(): ' + error.toString());
    return JSON.stringify({ success: false, error: 'Server-side exception during data fetch.' });
  }
}

// =====================================
// USER MANAGEMENT & AUTHENTICATION
// =====================================

/**
 * Get current user information
 */
function getCurrentUser() {
  try {
    const user = Session.getActiveUser();
    const email = user.getEmail();
    if (!email) {
      return { isSignedIn: false };
    }
    const name = email.split('@')[0]
      .split('.')
      .map(part => part.charAt(0).toUpperCase() + part.slice(1))
      .join(' ');
    const avatar = getOrAssignUserAvatar(email);
    return { isSignedIn: true, email: email, name: name, avatar: avatar };
  } catch (error) {
    console.error('Error getting current user:', error);
    return { isSignedIn: false };
  }
}

/**
 * Track user session for online presence
 */
function trackUserSession(user, action = 'heartbeat') {
  try {
    if (!user) {
      user = getCurrentUser();
    }
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('OnlineUsers');
    if (!sheet) {
      console.error('OnlineUsers sheet not found');
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    const now = new Date();
    cleanupExpiredSessions(sheet);
    if (action === 'logout') {
      removeUserSession(sheet, user.email);
      return JSON.stringify({ success: true });
    }
    const data = sheet.getDataRange().getValues();
    let sessionRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === user.email) {
        sessionRow = i + 1;
        break;
      }
    }
    if (action === 'login' || sessionRow === -1) {
      const sessionData = [generateId('SES'), user.email, user.name, now, now, user.avatar, 'Online'];
      if (sessionRow > 0) {
        sheet.getRange(sessionRow, 1, 1, sessionData.length).setValues([sessionData]);
      } else {
        sheet.appendRow(sessionData);
      }
    } else if (sessionRow > 0) {
      sheet.getRange(sessionRow, 5).setValue(now);
      sheet.getRange(sessionRow, 7).setValue('Online');
    }
    return JSON.stringify({ success: true });
  } catch (error) {
    console.error('Error tracking user session:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

/**
 * Get currently online users
 */
function getOnlineUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('OnlineUsers');
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify({ success: true, users: [] });
    }
    cleanupExpiredSessions(sheet);
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === 'Online') {
        users.push({
          email: data[i][1], name: data[i][2], loginTime: data[i][3],
          lastActivity: data[i][4], avatar: data[i][5], status: data[i][6]
        });
      }
    }
    users.sort((a, b) => new Date(b.lastActivity) - new Date(a.lastActivity));
    return JSON.stringify({ success: true, users: users });
  } catch (error) {
    console.error('Error getting online users:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// =====================================
// DOCUMENT MANAGEMENT (CRUD)
// =====================================

function getDocuments(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, documents: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const documents = [];
    for (let i = 1; i < data.length; i++) {
      const doc = {};
      headers.forEach((header, index) => { doc[header] = data[i][index]; });
      if (passesFilters(doc, filters)) {
        documents.push(doc);
      }
    }
    documents.sort(getSortFunction(filters.sortBy || 'date_desc'));
    return { success: true, documents: documents };
  } catch (error) {
    console.error('Error getting documents:', error);
    return { success: false, error: error.toString() };
  }
}

function addDocument(documentData) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    if (!documentData.DocumentName || !documentData.GoogleDriveURL || !documentData.Category) {
      return JSON.stringify({ success: false, error: 'Missing required fields' });
    }
    if (isDuplicateURL(documentData.GoogleDriveURL)) {
      return JSON.stringify({ success: false, error: 'A document with this URL already exists' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    const docId = generateId('DOC');
    const now = new Date();
    const fileType = detectFileType(documentData.GoogleDriveURL);
    const newDocument = [docId, documentData.DocumentName, documentData.GoogleDriveURL, documentData.Description || '', documentData.Category, fileType, user.email, documentData.Tags || '', now, now, 'Active'];
    sheet.appendRow(newDocument);
    
    logActivity(user, 'Created Document', docId, `Created "${documentData.DocumentName}"`);
    updateCategoryCount(documentData.Category, 1);
    if (documentData.Tags) updateTagCounts(documentData.Tags, 1);
    
    return JSON.stringify({ success: true, docId: docId, message: 'Document added successfully' });
  } catch (error) {
    console.error('Error adding document:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function updateDocument(docId, updates) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rowIndex = data.findIndex(row => row[0] === docId);
    
    if (rowIndex === -1) {
      return JSON.stringify({ success: false, error: 'Document not found' });
    }
    
    const changes = [];
    const oldCategory = data[rowIndex][headers.indexOf('Category')];
    const oldTags = data[rowIndex][headers.indexOf('Tags')];
    
    Object.keys(updates).forEach(field => {
      const colIndex = headers.indexOf(field);
      if (colIndex > -1 && updates[field] !== undefined && data[rowIndex][colIndex] !== updates[field]) {
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(updates[field]);
        changes.push(`${field} updated`);
      }
    });
    
    sheet.getRange(rowIndex + 1, headers.indexOf('LastModified') + 1).setValue(new Date());
    
    if (updates.Category && updates.Category !== oldCategory) {
      updateCategoryCount(oldCategory, -1);
      updateCategoryCount(updates.Category, 1);
    }
    if (updates.Tags !== oldTags) {
      updateTagCounts(oldTags, -1);
      updateTagCounts(updates.Tags, 1);
    }
    
    if (changes.length > 0) logActivity(user, 'Updated Document', docId, changes.join(', '));
    
    return JSON.stringify({ success: true, message: 'Document updated successfully' });
  } catch (error) {
    console.error('Error updating document:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function archiveDocument(docId) {
  try {
    const user = getCurrentUser();
    const result = JSON.parse(updateDocument(docId, { Status: 'Archived' }));
    if (result.success) logActivity(user, 'Archived Document', docId, 'Document archived');
    return JSON.stringify(result);
  } catch (error) {
    console.error('Error archiving document:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function restoreDocument(docId) {
  try {
    const user = getCurrentUser();
    const result = JSON.parse(updateDocument(docId, { Status: 'Active' }));
    if (result.success) logActivity(user, 'Restored Document', docId, 'Document restored');
    return JSON.stringify(result);
  } catch (error) {
    console.error('Error restoring document:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function deleteDocument(docId) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === docId);

    if (rowIndex === -1) {
      return JSON.stringify({ success: false, error: 'Document not found' });
    }
    
    const docName = data[rowIndex][1];
    const category = data[rowIndex][4];
    const tags = data[rowIndex][7];
    
    sheet.deleteRow(rowIndex + 1);
    
    updateCategoryCount(category, -1);
    if (tags) updateTagCounts(tags, -1);
    removeFromAllFavorites(docId);
    logActivity(user, 'Deleted Document', docId, `Permanently deleted "${docName}"`);
    
    return JSON.stringify({ success: true, message: 'Document deleted successfully' });
  } catch (error) {
    console.error('Error deleting document:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// =====================================
// BULK OPERATIONS
// =====================================

function bulkArchiveDocuments(docIds) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    let successCount = 0;
    docIds.forEach(docId => {
      const result = JSON.parse(updateDocument(docId, { Status: 'Archived' }));
      if (result.success) {
        logActivity(user, 'Archived Document', docId, 'Bulk operation');
        successCount++;
      }
    });
    return JSON.stringify({ success: true, summary: `${successCount} of ${docIds.length} documents archived.` });
  } catch (error) {
    console.error('Error in bulk archive:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function bulkRestoreDocuments(docIds) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    let successCount = 0;
    docIds.forEach(docId => {
      const result = JSON.parse(updateDocument(docId, { Status: 'Active' }));
      if (result.success) {
        logActivity(user, 'Restored Document', docId, 'Bulk operation');
        successCount++;
      }
    });
    return JSON.stringify({ success: true, summary: `${successCount} of ${docIds.length} documents restored.` });
  } catch (error) {
    console.error('Error in bulk restore:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// =====================================
// CATEGORIES & TAGS
// =====================================

function getCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Categories');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, categories: [] };
    const data = sheet.getDataRange().getValues();
    const categories = data.slice(1).filter(row => row[4] === true).map(row => ({
      CategoryID: row[0], CategoryName: row[1], CreatedBy: row[2], DateCreated: row[3], DocumentCount: row[5] || 0
    }));
    categories.sort((a, b) => a.CategoryName.localeCompare(b.CategoryName));
    return { success: true, categories: categories };
  } catch (error) {
    console.error('Error getting categories:', error);
    return { success: false, error: error.toString() };
  }
}

function addCategory(categoryName) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Categories');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    // Check if category already exists
    const data = sheet.getDataRange().getValues();
    const existingCategory = data.find(row => row[1] === categoryName);
    if (existingCategory) {
      return JSON.stringify({ success: false, error: 'Category already exists' });
    }
    
    const categoryId = generateId('CAT');
    const now = new Date();
    sheet.appendRow([categoryId, categoryName, user.email, now, true, 0]);
    
    logActivity(user, 'Created Category', categoryId, `Created category "${categoryName}"`);
    
    return JSON.stringify({ success: true, categoryId: categoryId, message: 'Category added successfully' });
  } catch (error) {
    console.error('Error adding category:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function getTags() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Tags');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, tags: [] };
    const data = sheet.getDataRange().getValues();
    const tags = data.slice(1).map(row => ({
      TagID: row[0], TagName: row[1], CreatedBy: row[2], DateCreated: row[3], UsageCount: row[4] || 0
    }));
    tags.sort((a, b) => b.UsageCount - a.UsageCount);
    return { success: true, tags: tags };
  } catch (error) {
    console.error('Error getting tags:', error);
    return { success: false, error: error.toString() };
  }
}

// =====================================
// FAVORITES
// =====================================

function getUserFavorites() {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) return { success: false, error: 'User not authenticated' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('UserFavorites');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, favoriteIds: [] };
    
    const data = sheet.getDataRange().getValues();
    const favoriteIds = data.slice(1)
      .filter(row => row[1] === user.email)
      .map(row => row[2]);
      
    return { success: true, favoriteIds: favoriteIds };
  } catch (error) {
    console.error('Error getting user favorites:', error);
    return { success: false, error: error.toString() };
  }
}

function toggleFavorite(docId) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('UserFavorites');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[1] === user.email && row[2] === docId);

    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
      logActivity(user, 'Removed Favorite', docId, 'Removed from favorites');
      return JSON.stringify({ success: true, favorited: false });
    } else {
      sheet.appendRow([generateId('FAV'), user.email, docId, new Date()]);
      logActivity(user, 'Added Favorite', docId, 'Added to favorites');
      return JSON.stringify({ success: true, favorited: true });
    }
  } catch (error) {
    console.error('Error toggling favorite:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// =====================================
// ANALYTICS & ACTIVITY
// =====================================

function recordDocumentView(docId) {
  try {
    const user = getCurrentUser();
    if (!user.isSignedIn) {
      return JSON.stringify({ success: false, error: 'User not authenticated' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Analytics');
    if (!sheet) {
      return JSON.stringify({ success: false, error: 'Database not initialized' });
    }
    
    // Get document name for the view record
    const docSheet = ss.getSheetByName('Documents');
    const docData = docSheet.getDataRange().getValues();
    const docRow = docData.find(row => row[0] === docId);
    const docName = docRow ? docRow[1] : 'Unknown Document';
    
    const viewId = generateId('VIEW');
    const now = new Date();
    sheet.appendRow([viewId, docId, user.email, now, 'Web Browser', docName]);
    
    return JSON.stringify({ success: true, message: 'View recorded' });
  } catch (error) {
    console.error('Error recording document view:', error);
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

function getAnalyticsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Analytics');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, analytics: { totalViews: 0, topDocuments: [], topCategories: {}, recentViews: [], viewsByDay: {} } };
    }
    const totalViews = sheet.getLastRow() - 1;
    return { success: true, analytics: { totalViews: totalViews, topDocuments: [], topCategories: {}, recentViews: [], viewsByDay: {} } };
  } catch(error) {
    console.error('Error getting analytics:', error);
    return { success: false, analytics: {} };
  }
}

function getRecentActivity(limit = CONFIG.recentActivityLimit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ActivityLog');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, activities: [] };
    
    const data = sheet.getDataRange().getValues();
    const activities = data.slice(Math.max(1, data.length - limit)).reverse().map(row => ({
      ActivityID: row[0], UserEmail: row[1], UserName: row[2], Action: row[3],
      DocID: row[4], Details: row[5], Timestamp: row[6]
    }));
    return { success: true, activities: activities };
  } catch (error) {
    console.error('Error getting recent activity:', error);
    return { success: false, error: error.toString() };
  }
}

// =====================================
// UTILITY FUNCTIONS
// =====================================

/**
 * Generate unique IDs for database records
 */
function generateId(prefix = 'ID') {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substr(2, 5);
  return `${prefix}_${timestamp}_${random}`.toUpperCase();
}

/**
 * Detect file type from Google Drive URL
 */
function detectFileType(url) {
  if (!url) return 'Other';
  if (url.includes('docs.google.com/document')) return 'Google Doc';
  if (url.includes('docs.google.com/spreadsheets')) return 'Google Sheet';
  if (url.includes('docs.google.com/presentation')) return 'Google Slides';
  if (url.includes('docs.google.com/forms')) return 'Google Form';
  if (url.includes('drive.google.com/file')) {
    // Try to detect from file extension or other indicators
    if (url.includes('.pdf')) return 'PDF';
    if (url.includes('.jpg') || url.includes('.png') || url.includes('.gif')) return 'Image';
  }
  if (url.includes('sites.google.com')) return 'Google Site';
  if (url.includes('lookerstudio.google.com')) return 'Looker Studio';
  if (url.includes('tableau.com')) return 'Tableau';
  return 'Website';
}

/**
 * Remove document from all user favorites
 */
function removeFromAllFavorites(docId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('UserFavorites');
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === docId) {
        rowsToDelete.push(i + 1);
      }
    }
    
    // Delete rows in reverse order to maintain row indices
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  } catch (error) {
    console.error('Error removing from favorites:', error);
  }
}

function getOrAssignUserAvatar(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('OnlineUsers');
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === email && data[i][5]) {
          return data[i][5];
        }
      }
    }
    const hash = email.split('').reduce((acc, char) => (((acc << 5) - acc) + char.charCodeAt(0)) | 0, 0);
    return CONFIG.animalAvatars[Math.abs(hash) % CONFIG.animalAvatars.length];
  } catch (error) {
    return '🦄';
  }
}

function isDuplicateURL(url) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet || sheet.getLastRow() <= 1) return false;
    const urls = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    return urls.some(row => row[0] === url);
  } catch (error) {
    console.error('Error checking duplicate URL:', error);
    return false;
  }
}

function passesFilters(doc, filters) {
  if (filters.status && filters.status.length > 0 && !filters.status.includes(doc.Status)) return false;
  if (filters.categories && filters.categories.length > 0 && !filters.categories.includes(doc.Category)) return false;
  if (filters.fileTypes && filters.fileTypes.length > 0 && !filters.fileTypes.includes(doc.FileType)) return false;
  return true;
}

function getSortFunction(sortBy) {
  switch (sortBy) {
    case 'name_asc': return (a, b) => a.DocumentName.localeCompare(b.DocumentName);
    case 'name_desc': return (a, b) => b.DocumentName.localeCompare(a.DocumentName);
    case 'date_asc': return (a, b) => new Date(a.DateAdded) - new Date(b.DateAdded);
    default: return (a, b) => new Date(b.LastModified) - new Date(a.LastModified);
  }
}

function logActivity(user, action, docId, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ActivityLog');
    if (sheet) {
      sheet.appendRow([generateId('ACT'), user.email, user.name, action, docId || '', details || '', new Date()]);
    }
  } catch (error) {
    console.error('Error logging activity:', error);
  }
}

function updateCategoryCount(categoryName, delta) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Categories');
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[1] === categoryName);
    if (rowIndex > 0) {
      const currentCount = parseInt(data[rowIndex][5]) || 0;
      sheet.getRange(rowIndex + 1, 6).setValue(Math.max(0, currentCount + delta));
    }
  } catch (error) {
    console.error(`Error updating count for category ${categoryName}:`, error);
  }
}

function updateTagCounts(tagsString, delta) {
  if (!tagsString) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Tags');
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const tags = tagsString.split(',').map(t => t.trim());
    tags.forEach(tag => {
      const rowIndex = data.findIndex(row => row[1] === tag);
      if (rowIndex > 0) {
        const currentCount = parseInt(data[rowIndex][4]) || 0;
        sheet.getRange(rowIndex + 1, 5).setValue(Math.max(0, currentCount + delta));
      }
    });
  } catch (error) {
    console.error(`Error updating counts for tags "${tagsString}":`, error);
  }
}

function cleanupExpiredSessions(sheet) {
  try {
    if (!sheet || sheet.getLastRow() <= 1) return;
    const cutoffTime = new Date(Date.now() - CONFIG.sessionTimeout * 60 * 1000);
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const lastActivity = new Date(data[i][4]);
      if (lastActivity < cutoffTime) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  } catch (error) {
    console.error('Error cleaning up expired sessions:', error);
  }
}

function removeUserSession(sheet, email) {
  try {
    if (!sheet || sheet.getLastRow() <= 1) return;
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[1] === email);
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
    }
  } catch (error) {
    console.error(`Error removing session for ${email}:`, error);
  }
}

function createAuthRequiredPage() {
  const html = `
    <!DOCTYPE html><html><head><title>Authentication Required</title><style>
    body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background-color:#f3f4f6;}
    .c{text-align:center;padding:48px;background:white;border-radius:8px;box-shadow:0 4px 6px rgba(0,0,0,0.1);}
    h1{color:#1e3a8a;margin-bottom:16px;} p{color:#6b7280;margin-bottom:24px;}
    .b{background:#1e3a8a;color:white;padding:12px 24px;border:none;border-radius:6px;font-size:16px;cursor:pointer;}
    </style></head><body><div class="c"><h1>🔒 Authentication Required</h1>
    <p>Please sign in to access the Document Center.</p><button class="b" onclick="window.location.reload()">Sign In</button>
    </div></body></html>`;
  return HtmlService.createHtmlOutput(html);
}

function createErrorPage(msg) {
  const html = `
    <!DOCTYPE html><html><head><title>Error</title><style>
    body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background-color:#f3f4f6;}
    .c{text-align:center;padding:48px;background:white;border-radius:8px;box-shadow:0 4px 6px rgba(0,0,0,0.1);max-width:500px;}
    h1{color:#ef4444;margin-bottom:16px;} p{color:#6b7280;margin-bottom:24px;}
    .d{background:#fef2f2;color:#991b1b;padding:12px;border-radius:6px;margin-bottom:24px;font-family:monospace;}
    .b{background:#1e3a8a;color:white;padding:12px 24px;border:none;border-radius:6px;font-size:16px;cursor:pointer;}
    </style></head><body><div class="c"><h1>⚠️ Something went wrong</h1>
    <p>We encountered an error.</p><div class="d">${msg}</div><button class="b" onclick="window.location.reload()">Try Again</button>
    </div></body></html>`;
  return HtmlService.createHtmlOutput(html);
}
