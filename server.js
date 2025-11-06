const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { promisify } = require('util');
const writeFile = promisify(fs.writeFile);
const readFile = promisify(fs.readFile);
const mkdir = promisify(fs.mkdir);
const unlink = promisify(fs.unlink);
const readdir = promisify(fs.readdir);

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(fileUpload());
app.use('/uploads', express.static('uploads'));

const PORT = 3000;
const EXCEL_FILE = 'inventory.xlsx';
const UPLOAD_DIR = 'uploads';
const USERS_FILE = 'users.json';
const ADMIN_USERS_FILE = 'admin_users.json';

// Ensure upload directory exists
async function ensureUploadDir() {
    try {
        await mkdir(UPLOAD_DIR, { recursive: true });
    } catch (err) {
        if (err.code !== 'EEXIST') throw err;
    }
}

// Initialize users files
async function initializeUsersFiles() {
    try {
        if (!fs.existsSync(USERS_FILE)) {
            await writeFile(USERS_FILE, JSON.stringify([], null, 2));
        }
        if (!fs.existsSync(ADMIN_USERS_FILE)) {
            await writeFile(ADMIN_USERS_FILE, JSON.stringify([], null, 2));
        }
    } catch (error) {
        console.error("Error initializing user files:", error);
    }
}

// Load users from file
async function loadUsers(isAdmin = false) {
    try {
        const file = isAdmin ? ADMIN_USERS_FILE : USERS_FILE;
        const data = await readFile(file, 'utf8');
        return JSON.parse(data);
    } catch (error) {
        console.error("Error loading users:", error);
        return [];
    }
}

// Save users to file
async function saveUsers(users, isAdmin = false) {
    try {
        const file = isAdmin ? ADMIN_USERS_FILE : USERS_FILE;
        await writeFile(file, JSON.stringify(users, null, 2));
        return true;
    } catch (error) {
        console.error("Error saving users:", error);
        return false;
    }
}

// Check if username exists
async function usernameExists(username, isAdmin = false) {
    const users = await loadUsers(isAdmin);
    return users.some(user => user.username.toLowerCase() === username.toLowerCase());
}

// Find user by username and password
async function findUser(username, password, isAdmin = false) {
    const users = await loadUsers(isAdmin);
    return users.find(user => 
        user.username.toLowerCase() === username.toLowerCase() && 
        user.password === password
    );
}

// Register new user
async function registerUser(username, password, isAdmin = false) {
    try {
        const users = await loadUsers(isAdmin);
        
        if (await usernameExists(username, isAdmin)) {
            return { success: false, message: "Username already exists" };
        }
        
        const newUser = {
            id: `${isAdmin ? 'ADMIN' : 'USER'}-${Date.now()}`,
            username: username,
            password: password,
            createdAt: new Date().toISOString(),
            role: isAdmin ? 'admin' : 'user'
        };
        
        users.push(newUser);
        const saved = await saveUsers(users, isAdmin);
        
        if (saved) {
            return { 
                success: true, 
                message: "User registered successfully",
                user: { id: newUser.id, username: newUser.username, role: newUser.role }
            };
        } else {
            return { success: false, message: "Failed to save user" };
        }
    } catch (error) {
        console.error("Error registering user:", error);
        return { success: false, message: "Registration failed" };
    }
}

// Load data from Excel (or create new file if it doesn't exist)
function loadExcelSync() {
    try {
        if (!fs.existsSync(EXCEL_FILE)) {
            const defaultData = [];
            const newWB = xlsx.utils.book_new();
            const newWS = xlsx.utils.json_to_sheet(defaultData);
            xlsx.utils.book_append_sheet(newWB, newWS, "Inventory");
            xlsx.writeFile(newWB, EXCEL_FILE);
            return defaultData;
        }

        const wb = xlsx.readFile(EXCEL_FILE);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(ws);
        
        return data.filter(item => 
            (item['Component ID'] && item['Component ID'].toString().trim() !== '') ||
            (item['Part No'] && item['Part No'].toString().trim() !== '')
        );
    } catch (error) {
        console.error("Error loading Excel:", error);
        return [];
    }
}

// Save data to Excel - SYNCHRONOUS VERSION
function saveToExcelSync(data) {
    try {
        console.log('Attempting to save data to Excel...');
        console.log('Number of items:', data.length);
        
        // Create new workbook
        const newWB = xlsx.utils.book_new();
        const newWS = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(newWB, newWS, "Inventory");
        
        // Write file synchronously
        xlsx.writeFile(newWB, EXCEL_FILE);
        
        console.log('Excel file saved successfully');
        
        // Verify the save
        const verifyWB = xlsx.readFile(EXCEL_FILE);
        const verifyWS = verifyWB.Sheets[verifyWB.SheetNames[0]];
        const verifyData = xlsx.utils.sheet_to_json(verifyWS);
        console.log('Verification: Items in file:', verifyData.length);
        
        return true;
    } catch (error) {
        console.error("Error saving Excel:", error);
        console.error("Error details:", error.message);
        return false;
    }
}

// Generate unique component ID
function generateComponentId(existingData) {
    const existingIds = existingData
        .filter(item => item['Component ID'])
        .map(item => {
            const match = item['Component ID'].toString().match(/CMP-(\d+)/);
            return match ? parseInt(match[1]) : 0;
        });
    
    const maxId = Math.max(0, ...existingIds);
    return `CMP-${(maxId + 1).toString().padStart(3, '0')}`;
}

// API Endpoints

// Regular user signup
app.post('/api/signup', async (req, res) => {
    const { username, password } = req.body;
    
    if (!username || !password) {
        return res.status(400).json({ success: false, message: "Username and password are required" });
    }
    
    if (username.length < 3) {
        return res.status(400).json({ success: false, message: "Username must be at least 3 characters long" });
    }
    
    if (password.length < 4) {
        return res.status(400).json({ success: false, message: "Password must be at least 4 characters long" });
    }
    
    const result = await registerUser(username, password, false);
    
    if (result.success) {
        res.json({ 
            success: true, 
            message: "Signup successful! Please login.",
            user: result.user
        });
    } else {
        res.status(400).json(result);
    }
});

// Regular user login
app.post('/api/login', async (req, res) => {
    const { username, password } = req.body;
    
    if (!username || !password) {
        return res.status(400).json({ success: false, message: "Username and password are required" });
    }
    
    const user = await findUser(username, password, false);
    
    if (user) {
        res.json({ 
            success: true,
            message: "Login successful!",
            user: { 
                id: user.id,
                name: user.username,
                role: 'user',
                token: 'user-token-' + Math.random().toString(36).substr(2)
            } 
        });
    } else {
        res.status(401).json({ 
            success: false, 
            message: "Invalid username or password" 
        });
    }
});

// Admin signup
app.post('/api/admin/signup', async (req, res) => {
    const { username, password } = req.body;
    
    if (!username || !password) {
        return res.status(400).json({ success: false, message: "Username and password are required" });
    }
    
    if (username.length < 3) {
        return res.status(400).json({ success: false, message: "Username must be at least 3 characters long" });
    }
    
    if (password.length < 4) {
        return res.status(400).json({ success: false, message: "Password must be at least 4 characters long" });
    }
    
    const result = await registerUser(username, password, true);
    
    if (result.success) {
        res.json({ 
            success: true, 
            message: "Admin signup successful! Please login.",
            user: result.user
        });
    } else {
        res.status(400).json(result);
    }
});

// Admin login endpoint
app.post('/api/admin/login', async (req, res) => {
    const { username, password } = req.body;
    
    if (!username || !password) {
        return res.status(400).json({ success: false, message: "Username and password are required" });
    }
    
    const user = await findUser(username, password, true);
    
    if (user) {
        res.json({ 
            success: true,
            message: "Admin login successful!",
            user: { 
                id: user.id,
                name: user.username,
                role: 'admin',
                token: 'admin-token-' + Math.random().toString(36).substr(2)
            } 
        });
    } else {
        res.status(401).json({ 
            success: false, 
            message: "Invalid admin credentials" 
        });
    }
});

// Get inventory data
app.get('/api/inventory', async (req, res) => {
    try {
        const data = loadExcelSync();
        res.setHeader('Cache-Control', 'no-store');
        res.json(data);
    } catch (error) {
        console.error("Error fetching inventory:", error);
        res.status(500).json({ error: "Failed to load inventory" });
    }
});

// Get pending requests for admin
app.get('/api/requests/pending', async (req, res) => {
    try {
        const data = loadExcelSync();
        
        // Group by Issue No or Storage No
        const groupedRequests = {};
        
        data.forEach(item => {
            if (item.Status && item.Status.toString().toLowerCase() === 'pending') {
                const groupKey = item['Issue No'] || item['Storage No'];
                
                if (groupKey) {
                    if (!groupedRequests[groupKey]) {
                        groupedRequests[groupKey] = { ...item };
                    }
                }
            }
        });
        
        const pendingRequests = Object.values(groupedRequests);
        console.log('Pending requests found:', pendingRequests.length);
        res.json(pendingRequests);
    } catch (error) {
        console.error("Error fetching pending requests:", error);
        res.status(500).json({ error: "Failed to load pending requests" });
    }
});

// Issue components endpoint with file upload support
app.post('/api/issue', async (req, res) => {
    try {
        await ensureUploadDir();
        const currentData = loadExcelSync();
        
        // Parse components from form data
        const components = JSON.parse(req.body.components);
        const { issueNo, issueDate, requestText, issueTo, issueFor, systemManager, submittedBy } = req.body;
        
        if (!issueNo || !issueDate || !issueTo || !components || components.length === 0) {
            return res.status(400).json({ error: "Missing required fields" });
        }
        
        const newItems = [];
        
        for (let i = 0; i < components.length; i++) {
            const component = components[i];
            
            // Handle PDF file upload if exists
            let pdfFileName = null;
            if (component.hasPdf && req.files && req.files[`soPdf_${component.pdfIndex}`]) {
                const pdfFile = req.files[`soPdf_${component.pdfIndex}`];
                const timestamp = Date.now();
                const sanitizedPartNo = component.partNo.replace(/[^a-zA-Z0-9]/g, '_');
                pdfFileName = `SO_${issueNo}_${sanitizedPartNo}_${timestamp}.pdf`;
                const pdfPath = path.join(UPLOAD_DIR, pdfFileName);
                
                await pdfFile.mv(pdfPath);
                console.log(`PDF saved: ${pdfFileName}`);
            }
            
            const newItem = {
                "Component ID": generateComponentId([...currentData, ...newItems]),
                "Name": component.partDescription || "Issued Component",
                "Part No": component.partNo,
                "Part Description": component.partDescription,
                "Type": "Issued Component",
                "Status": "Pending",
                "Date": new Date().toISOString().split('T')[0],
                "Issued To": issueTo,
                "Issue No": issueNo,
                "Issue Date": component.issueDate || issueDate,
                "Request Text": requestText,
                "Issue For": issueFor,
                "System Manager": systemManager,
                "Serial No": component.serialNo,
                "S.No as per SO": component.snoSO,
                "Manufacturer": component.manufacturer,
                "Quality Grade": component.qualityGrade,
                "Sub System": component.subSystem,
                "Quantity Each": component.quantityEach,
                "Total Quantity": component.totalQuantity,
                "SO No": component.soNo,
                "SO PDF": pdfFileName ? `/uploads/${pdfFileName}` : null,
                "Submitted By": submittedBy
            };
            newItems.push(newItem);
        }
        
        const updatedData = [...currentData, ...newItems];
        const success = saveToExcelSync(updatedData);
        
        if (success) {
            res.json({ 
                success: true, 
                message: `Successfully submitted ${components.length} components for approval`,
                data: updatedData 
            });
        } else {
            res.status(500).json({ error: "Failed to save issue data" });
        }
    } catch (error) {
        console.error("Error processing issue:", error);
        res.status(500).json({ error: "Failed to process component issue: " + error.message });
    }
});

// Storage components endpoint
app.post('/api/storage', async (req, res) => {
    try {
        const currentData = loadExcelSync();
        const { storageNo, storageDate, soNumber, systemManager, components } = req.body;
        
        if (!storageNo || !storageDate || !soNumber || !components || components.length === 0) {
            return res.status(400).json({ error: "Missing required fields" });
        }
        
        const newItems = [];
        
        for (const component of components) {
            const newItem = {
                "Component ID": generateComponentId([...currentData, ...newItems]),
                "Name": component.partDescription || "Stored Component",
                "Part No": component.partNo,
                "Part Description": component.partDescription,
                "Type": "Stored Component",
                "Status": "Pending",
                "Date": new Date().toISOString().split('T')[0],
                "Storage No": storageNo,
                "Storage Date": storageDate,
                "SO Number": soNumber,
                "System Manager": systemManager,
                "Serial No": component.serialNo,
                "S.No as per PO": component.snoPO,
                "Grade": component.grade,
                "Storage Quantity": component.quantity,
                "Storage Temperature": component.storageTemp,
                "Relative Humidity": component.relativeHumidity,
                "Storage Data": component.storageData,
                "Delivery Date": component.deliveryDate,
                "SO No": soNumber
            };
            newItems.push(newItem);
        }
        
        const updatedData = [...currentData, ...newItems];
        const success = saveToExcelSync(updatedData);
        
        if (success) {
            res.json({ 
                success: true, 
                message: `Successfully submitted ${components.length} components for storage approval`,
                data: updatedData 
            });
        } else {
            res.status(500).json({ error: "Failed to save storage data" });
        }
    } catch (error) {
        console.error("Error processing storage:", error);
        res.status(500).json({ error: "Failed to process component storage" });
    }
});

// Approve request endpoint - COMPLETELY FIXED
app.post('/api/requests/approve', async (req, res) => {
    try {
        const { requestId, approvalData } = req.body;
        console.log('\n=== APPROVAL REQUEST ===');
        console.log('Request ID:', requestId);
        console.log('Approved By:', approvalData.approvedBy);
        
        const currentData = loadExcelSync();
        console.log('Total items in Excel:', currentData.length);
        
        // Find all items with matching Issue No or Storage No or Component ID
        let matchingIndices = [];
        let updatedCount = 0;
        
        for (let i = 0; i < currentData.length; i++) {
            const item = currentData[i];
            const issueNo = item['Issue No'] ? item['Issue No'].toString() : null;
            const storageNo = item['Storage No'] ? item['Storage No'].toString() : null;
            const componentId = item['Component ID'] ? item['Component ID'].toString() : null;
            const reqId = requestId.toString();
            
            if (issueNo === reqId || storageNo === reqId || componentId === reqId) {
                matchingIndices.push(i);
                console.log(`Match found at index ${i}:`, item['Component ID']);
            }
        }
        
        console.log('Matching items found:', matchingIndices.length);
        
        if (matchingIndices.length === 0) {
            console.log('ERROR: No matching items found');
            return res.status(404).json({ 
                success: false,
                error: "Request not found in database" 
            });
        }
        
        // Update all matching items
        matchingIndices.forEach(index => {
            currentData[index]['Status'] = 'Approved';
            currentData[index]['Approved By'] = approvalData.approvedBy;
            currentData[index]['Approval Date'] = new Date().toISOString().split('T')[0];
            currentData[index]['Approval Signature'] = approvalData.signature;
            updatedCount++;
            console.log(`Updated item ${updatedCount}:`, currentData[index]['Component ID']);
        });
        
        console.log('Total items updated:', updatedCount);
        console.log('Attempting to save to Excel...');
        
        const success = saveToExcelSync(currentData);
        
        if (success) {
            console.log('SUCCESS: Approval saved to Excel');
            res.json({ 
                success: true, 
                message: `Request approved successfully! ${updatedCount} item(s) updated.`,
                itemsUpdated: updatedCount
            });
        } else {
            console.log('ERROR: Failed to save to Excel');
            res.status(500).json({ 
                success: false,
                error: "Failed to save approval to Excel file" 
            });
        }
    } catch (error) {
        console.error("ERROR in approval:", error);
        console.error("Stack trace:", error.stack);
        res.status(500).json({ 
            success: false,
            error: "Server error: " + error.message 
        });
    }
});

// Reject request endpoint - COMPLETELY FIXED
app.post('/api/requests/reject', async (req, res) => {
    try {
        const { requestId, rejectionReason } = req.body;
        console.log('\n=== REJECTION REQUEST ===');
        console.log('Request ID:', requestId);
        console.log('Rejection Reason:', rejectionReason);
        
        const currentData = loadExcelSync();
        console.log('Total items in Excel:', currentData.length);
        
        // Find all items with matching Issue No or Storage No or Component ID
        let matchingIndices = [];
        let updatedCount = 0;
        
        for (let i = 0; i < currentData.length; i++) {
            const item = currentData[i];
            const issueNo = item['Issue No'] ? item['Issue No'].toString() : null;
            const storageNo = item['Storage No'] ? item['Storage No'].toString() : null;
            const componentId = item['Component ID'] ? item['Component ID'].toString() : null;
            const reqId = requestId.toString();
            
            if (issueNo === reqId || storageNo === reqId || componentId === reqId) {
                matchingIndices.push(i);
                console.log(`Match found at index ${i}:`, item['Component ID']);
            }
        }
        
        console.log('Matching items found:', matchingIndices.length);
        
        if (matchingIndices.length === 0) {
            console.log('ERROR: No matching items found');
            return res.status(404).json({ 
                success: false,
                error: "Request not found in database" 
            });
        }
        
        // Update all matching items
        matchingIndices.forEach(index => {
            currentData[index]['Status'] = 'Rejected';
            currentData[index]['Rejection Reason'] = rejectionReason;
            currentData[index]['Rejection Date'] = new Date().toISOString().split('T')[0];
            updatedCount++;
            console.log(`Updated item ${updatedCount}:`, currentData[index]['Component ID']);
        });
        
        console.log('Total items updated:', updatedCount);
        console.log('Attempting to save to Excel...');
        
        const success = saveToExcelSync(currentData);
        
        if (success) {
            console.log('SUCCESS: Rejection saved to Excel');
            res.json({ 
                success: true, 
                message: `Request rejected successfully! ${updatedCount} item(s) updated.`,
                itemsUpdated: updatedCount
            });
        } else {
            console.log('ERROR: Failed to save to Excel');
            res.status(500).json({ 
                success: false,
                error: "Failed to save rejection to Excel file" 
            });
        }
    } catch (error) {
        console.error("ERROR in rejection:", error);
        console.error("Stack trace:", error.stack);
        res.status(500).json({ 
            success: false,
            error: "Server error: " + error.message 
        });
    }
});

// Archive File Endpoints

// Get all archived files
app.get('/api/archive/files', async (req, res) => {
    try {
        await ensureUploadDir();
        const files = await readdir(UPLOAD_DIR);
        
        const fileDetails = await Promise.all(files.map(async (filename) => {
            const filePath = path.join(UPLOAD_DIR, filename);
            const stats = await fs.promises.stat(filePath);
            
            return {
                id: `FILE-${filename.split('-')[0]}`,
                name: filename,
                path: `/uploads/${filename}`,
                size: stats.size,
                uploadDate: stats.birthtime.toISOString(),
                description: filename.includes('_desc_') ? 
                    filename.split('_desc_')[1].split('.')[0].replace(/-/g, ' ') : 
                    'No description'
            };
        }));
        
        res.json(fileDetails);
    } catch (error) {
        console.error("Error fetching archived files:", error);
        res.status(500).json({ error: "Failed to load archived files" });
    }
});

// Upload file to archive
app.post('/api/archive/upload', async (req, res) => {
    try {
        await ensureUploadDir();
        
        if (!req.files || !req.files.file) {
            return res.status(400).json({ error: "No file uploaded" });
        }
        
        const file = req.files.file;
        const description = req.body.description || '';
        const sanitizedDescription = description.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');
        const fileExt = path.extname(file.name);
        const fileNameBase = path.basename(file.name, fileExt);
        const timestamp = Date.now();
        
        const newFileName = `${timestamp}_${fileNameBase}_desc_${sanitizedDescription}${fileExt}`;
        const filePath = path.join(UPLOAD_DIR, newFileName);
        
        await file.mv(filePath);
        
        res.json({ 
            success: true, 
            message: "File uploaded successfully",
            filename: newFileName,
            path: `/uploads/${newFileName}`
        });
    } catch (error) {
        console.error("Error uploading file:", error);
        res.status(500).json({ error: "Failed to upload file" });
    }
});

// Delete file from archive
app.delete('/api/archive/files/:filename', async (req, res) => {
    try {
        const filename = req.params.filename;
        const filePath = path.join(UPLOAD_DIR, filename);
        
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: "File not found" });
        }
        
        await unlink(filePath);
        res.json({ success: true, message: "File deleted successfully" });
    } catch (error) {
        console.error("Error deleting file:", error);
        res.status(500).json({ error: "Failed to delete file" });
    }
});

// Rename file in archive
app.post('/api/archive/rename', async (req, res) => {
    try {
        const { oldName, newName } = req.body;
        const oldPath = path.join(UPLOAD_DIR, oldName);
        const newPath = path.join(UPLOAD_DIR, newName);
        
        if (!fs.existsSync(oldPath)) {
            return res.status(404).json({ error: "File not found" });
        }
        
        if (fs.existsSync(newPath)) {
            return res.status(400).json({ error: "A file with that name already exists" });
        }
        
        await fs.promises.rename(oldPath, newPath);
        res.json({ success: true, message: "File renamed successfully" });
    } catch (error) {
        console.error("Error renaming file:", error);
        res.status(500).json({ error: "Failed to rename file" });
    }
});

// Initialize server
async function startServer() {
    await ensureUploadDir();
    await initializeUsersFiles();
    
    app.listen(PORT, () => {
        console.log(`✓ Server running on http://localhost:${PORT}`);
        console.log(`✓ Excel file: ${EXCEL_FILE}`);
        console.log(`✓ User database: ${USERS_FILE}`);
        console.log(`✓ Admin database: ${ADMIN_USERS_FILE}`);
        console.log(`\nAPI Endpoints ready!`);
    });
}

startServer().catch(err => {
    console.error("Failed to start server:", err);
    process.exit(1);
});