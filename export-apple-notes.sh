#!/bin/bash

# Apple Notes to Markdown Export Script
# Uses obsidian-importer to export Apple Notes as Markdown files
# Requires macOS and Node.js

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Default values
OUTPUT_DIR=""
OMIT_FIRST_LINE=true
IMPORT_TRASHED=false
INCLUDE_HANDWRITING=false
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

print_help() {
    echo -e "${BLUE}Apple Notes to Markdown Export Script${NC}"
    echo ""
    echo "Usage: $0 [OPTIONS]"
    echo ""
    echo "Options:"
    echo "  -o, --output DIR         Output directory for exported notes (required)"
    echo "  --include-first-line     Include the first line in note content (default: omit)"
    echo "  --include-trashed        Include notes from 'Recently Deleted' folder"
    echo "  --include-handwriting    Include handwriting text from drawings"
    echo "  -h, --help              Show this help message"
    echo ""
    echo "Examples:"
    echo "  $0 -o ~/Documents/AppleNotesExport"
    echo "  $0 -o ./notes --include-trashed --include-handwriting"
    echo ""
    echo "Requirements:"
    echo "  - macOS (Apple Notes database access)"
    echo "  - Node.js and npm installed"
    echo "  - This script must be run from the obsidian-importer directory"
}

check_requirements() {
    echo -e "${BLUE}Checking requirements...${NC}"
    
    # Check if running on macOS
    if [[ "$OSTYPE" != "darwin"* ]]; then
        echo -e "${RED}Error: This script requires macOS to access Apple Notes database${NC}"
        exit 1
    fi
    
    # Check if Node.js is installed
    if ! command -v node &> /dev/null; then
        echo -e "${RED}Error: Node.js is not installed. Please install Node.js first.${NC}"
        exit 1
    fi
    
    # Check if npm is installed
    if ! command -v npm &> /dev/null; then
        echo -e "${RED}Error: npm is not installed. Please install npm first.${NC}"
        exit 1
    fi
    
    # Check if package.json exists (we're in the right directory)
    if [[ ! -f "$SCRIPT_DIR/package.json" ]]; then
        echo -e "${RED}Error: package.json not found. Please run this script from the obsidian-importer directory.${NC}"
        exit 1
    fi
    
    # Check if this is the obsidian-importer project
    if ! grep -q '"name": "obsidian-importer"' "$SCRIPT_DIR/package.json"; then
        echo -e "${RED}Error: This doesn't appear to be the obsidian-importer directory.${NC}"
        exit 1
    fi
    
    echo -e "${GREEN}✓ All requirements met${NC}"
}

install_dependencies() {
    echo -e "${BLUE}Installing dependencies...${NC}"
    
    cd "$SCRIPT_DIR"
    
    if [[ ! -d "node_modules" ]]; then
        echo -e "${YELLOW}Installing npm dependencies...${NC}"
        npm install
    else
        echo -e "${GREEN}✓ Dependencies already installed${NC}"
    fi
}

build_project() {
    echo -e "${BLUE}Building project...${NC}"
    
    cd "$SCRIPT_DIR"
    
    echo -e "${YELLOW}Building obsidian-importer...${NC}"
    npm run build
    
    echo -e "${GREEN}✓ Build completed${NC}"
}


export_notes() {
    local output_dir="$1"
    
    echo -e "${BLUE}Starting Apple Notes export...${NC}"
    
    # Create output directory if it doesn't exist
    mkdir -p "$output_dir"
    
    echo -e "${YELLOW}Output directory: $output_dir${NC}"
    
    # Create a Node.js script that directly processes Apple Notes database
    cat > "$SCRIPT_DIR/standalone-export.js" << 'EOF'
const fs = require('fs').promises;
const path = require('path');
const os = require('os');
const zlib = require('zlib');
const protobuf = require('protobufjs');

// Simple SQLite wrapper using sqlite3
let Database;
try {
    Database = require('sqlite3').Database;
} catch (e) {
    console.log('📦 Installing sqlite3 dependency...');
    require('child_process').execSync('npm install sqlite3', { stdio: 'inherit' });
    Database = require('sqlite3').Database;
}

class SimpleSQLiteWrapper {
    constructor(dbPath, options = {}) {
        this.db = new Database(dbPath, options.readonly ? 1 : 6);
    }
    
    all(query, params = []) {
        return new Promise((resolve, reject) => {
            this.db.all(query, params, (err, rows) => {
                if (err) reject(err);
                else resolve(rows || []);
            });
        });
    }
    
    get(query, params = []) {
        return new Promise((resolve, reject) => {
            this.db.get(query, params, (err, row) => {
                if (err) reject(err);
                else resolve(row);
            });
        });
    }
    
    close() {
        return new Promise((resolve) => {
            this.db.close(() => resolve());
        });
    }
}

class AppleNotesExporter {
    constructor(outputDir, options = {}) {
        this.outputDir = outputDir;
        this.omitFirstLine = options.omitFirstLine !== false;
        this.importTrashed = options.importTrashed === true;
        this.includeHandwriting = options.includeHandwriting === true;
        this.noteCount = 0;
        this.parsedNotes = 0;
        this.protobufRoot = null;
        this.attachmentCount = 0;
        this.dataPath = path.join(os.homedir(), 'Library/Group Containers/group.com.apple.notes');
        this.accounts = [];
        
        // Attachment type mappings
        this.ANAttachment = {
            Drawing: 'com.apple.paper',
            DrawingLegacy: 'com.apple.drawing',
            DrawingLegacy2: 'com.apple.drawing.2',
            Hashtag: 'com.apple.notes.inlinetextattachment.hashtag',
            Mention: 'com.apple.notes.inlinetextattachment.mention',
            InternalLink: 'com.apple.notes.inlinetextattachment.link',
            ModifiedScan: 'com.apple.paper.doc.scan',
            Scan: 'com.apple.notes.gallery',
            Table: 'com.apple.notes.table',
            UrlCard: 'public.url'
        };
    }
    
    async initProtobuf() {
        // Apple Notes protobuf descriptor
        const descriptor = {"nested":{"ciofecaforensics":{"nested":{"Color":{"fields":{"red":{"type":"float","id":1},"green":{"type":"float","id":2},"blue":{"type":"float","id":3},"alpha":{"type":"float","id":4}}},"AttachmentInfo":{"fields":{"attachmentIdentifier":{"type":"string","id":1},"typeUti":{"type":"string","id":2}}},"Font":{"fields":{"fontName":{"type":"string","id":1},"pointSize":{"type":"float","id":2},"fontHints":{"type":"int32","id":3}}},"ParagraphStyle":{"fields":{"styleType":{"type":"int32","id":1,"options":{"default":-1}},"alignment":{"type":"int32","id":2},"indentAmount":{"type":"int32","id":4},"checklist":{"type":"Checklist","id":5},"blockquote":{"type":"int32","id":8}}},"Checklist":{"fields":{"uuid":{"type":"bytes","id":1},"done":{"type":"int32","id":2}}},"DictionaryElement":{"fields":{"key":{"type":"ObjectID","id":1},"value":{"type":"ObjectID","id":2}}},"Dictionary":{"fields":{"element":{"rule":"repeated","type":"DictionaryElement","id":1,"options":{"packed":false}}}},"ObjectID":{"fields":{"unsignedIntegerValue":{"type":"uint64","id":2},"stringValue":{"type":"string","id":4},"objectIndex":{"type":"int32","id":6}}},"RegisterLatest":{"fields":{"contents":{"type":"ObjectID","id":2}}},"MapEntry":{"fields":{"key":{"type":"int32","id":1},"value":{"type":"ObjectID","id":2}}},"AttributeRun":{"fields":{"length":{"type":"int32","id":1},"paragraphStyle":{"type":"ParagraphStyle","id":2},"font":{"type":"Font","id":3},"fontWeight":{"type":"int32","id":5},"underlined":{"type":"int32","id":6},"strikethrough":{"type":"int32","id":7},"superscript":{"type":"int32","id":8},"link":{"type":"string","id":9},"color":{"type":"Color","id":10},"attachmentInfo":{"type":"AttachmentInfo","id":12}}},"NoteStoreProto":{"fields":{"document":{"type":"Document","id":2}}},"Document":{"fields":{"version":{"type":"int32","id":2},"note":{"type":"Note","id":3}}},"Note":{"fields":{"noteText":{"type":"string","id":2},"attributeRun":{"rule":"repeated","type":"AttributeRun","id":5,"options":{"packed":false}}}},"MergableDataProto":{"fields":{"mergableDataObject":{"type":"MergableDataObject","id":2}}},"MergableDataObject":{"fields":{"version":{"type":"int32","id":2},"mergeableDataObjectData":{"type":"MergeableDataObjectData","id":3}}},"MergeableDataObjectData":{"fields":{"mergeableDataObjectEntry":{"rule":"repeated","type":"MergeableDataObjectEntry","id":3,"options":{"packed":false}},"mergeableDataObjectKeyItem":{"rule":"repeated","type":"string","id":4},"mergeableDataObjectTypeItem":{"rule":"repeated","type":"string","id":5},"mergeableDataObjectUuidItem":{"rule":"repeated","type":"bytes","id":6}}},"MergeableDataObjectEntry":{"fields":{"registerLatest":{"type":"RegisterLatest","id":1},"list":{"type":"List","id":5},"dictionary":{"type":"Dictionary","id":6},"unknownMessage":{"type":"UnknownMergeableDataObjectEntryMessage","id":9},"note":{"type":"Note","id":10},"customMap":{"type":"MergeableDataObjectMap","id":13},"orderedSet":{"type":"OrderedSet","id":16}}},"UnknownMergeableDataObjectEntryMessage":{"fields":{"unknownEntry":{"type":"UnknownMergeableDataObjectEntryMessageEntry","id":1}}},"UnknownMergeableDataObjectEntryMessageEntry":{"fields":{"unknownInt1":{"type":"int32","id":1},"unknownInt2":{"type":"int64","id":2}}},"MergeableDataObjectMap":{"fields":{"type":{"type":"int32","id":1},"mapEntry":{"rule":"repeated","type":"MapEntry","id":3,"options":{"packed":false}}}},"OrderedSet":{"fields":{"ordering":{"type":"OrderedSetOrdering","id":1},"elements":{"type":"Dictionary","id":2}}},"OrderedSetOrdering":{"fields":{"array":{"type":"OrderedSetOrderingArray","id":1},"contents":{"type":"Dictionary","id":2}}},"OrderedSetOrderingArray":{"fields":{"contents":{"type":"Note","id":1},"attachment":{"rule":"repeated","type":"OrderedSetOrderingArrayAttachment","id":2,"options":{"packed":false}}}},"OrderedSetOrderingArrayAttachment":{"fields":{"index":{"type":"int32","id":1},"uuid":{"type":"bytes","id":2}}},"List":{"fields":{"listEntry":{"rule":"repeated","type":"ListEntry","id":1,"options":{"packed":false}}}},"ListEntry":{"fields":{"id":{"type":"ObjectID","id":2},"details":{"type":"ListEntryDetails","id":3},"additionalDetails":{"type":"ListEntryDetails","id":4}}},"ListEntryDetails":{"fields":{"listEntryDetailsKey":{"type":"ListEntryDetailsKey","id":1},"id":{"type":"ObjectID","id":2}}},"ListEntryDetailsKey":{"fields":{"listEntryDetailsTypeIndex":{"type":"int32","id":1},"listEntryDetailsKey":{"type":"int32","id":2}}}}}}};
        
        this.protobufRoot = protobuf.Root.fromJSON(descriptor);
    }
    
    async discoverAccounts(dataPath) {
        this.accounts = [];
        const accountsDir = path.join(dataPath, 'Accounts');
        
        try {
            const accountDirs = await fs.readdir(accountsDir);
            
            for (const accountDir of accountDirs) {
                if (accountDir.startsWith('.')) continue; // Skip hidden files
                
                const accountPath = path.join(accountsDir, accountDir);
                const stats = await fs.stat(accountPath);
                
                if (stats.isDirectory()) {
                    const mediaPath = path.join(accountPath, 'Media');
                    
                    // Check if Media directory exists
                    const hasMedia = await fs.access(mediaPath, fs.constants.F_OK).then(() => true).catch(() => false);
                    
                    this.accounts.push({
                        uuid: accountDir,
                        path: accountPath,
                        mediaPath: hasMedia ? mediaPath : null
                    });
                    
                    console.log(`📁 Found account: ${accountDir} (Media: ${hasMedia ? 'Yes' : 'No'})`);
                }
            }
        } catch (error) {
            console.warn('⚠️  Could not read accounts directory:', error.message);
            // Fallback to looking directly in the main data path
            this.accounts.push({
                uuid: 'default',
                path: dataPath,
                mediaPath: path.join(dataPath, 'Media')
            });
        }
    }
    
    async export() {
        console.log('🔍 Initializing protobuf decoder...');
        await this.initProtobuf();
        
        console.log('🔍 Locating Apple Notes database...');
        
        const dataPath = path.join(os.homedir(), 'Library/Group Containers/group.com.apple.notes');
        const originalDB = path.join(dataPath, 'NoteStore.sqlite');
        const clonedDB = path.join(os.tmpdir(), 'NoteStore.sqlite');
        
        // Find account directories for media files
        console.log('🔍 Discovering account structures...');
        await this.discoverAccounts(dataPath);
        
        try {
            // Check if database exists and is readable
            await fs.access(originalDB, fs.constants.R_OK);
            
            // Copy database files
            await fs.copyFile(originalDB, clonedDB);
            await fs.copyFile(originalDB + '-shm', clonedDB + '-shm').catch(() => {});
            await fs.copyFile(originalDB + '-wal', clonedDB + '-wal').catch(() => {});
            
            console.log('✅ Database copied successfully');
        } catch (error) {
            if (error.code === 'EPERM' || error.code === 'EACCES') {
                throw new Error(`Permission denied accessing Apple Notes database. Please:
1. Quit Apple Notes completely
2. Grant Terminal/iTerm full disk access in System Preferences > Security & Privacy > Privacy > Full Disk Access
3. Or run this script with sudo (not recommended)

Error details: ${error.message}`);
            }
            throw new Error(`Failed to access Apple Notes database: ${error.message}`);
        }
        
        // Open database
        console.log('🗂️  Opening database...');
        const database = new SimpleSQLiteWrapper(clonedDB, { readonly: true });
        
        try {
            // Get entity keys
            const keyRows = await database.all('SELECT z_ent, z_name FROM z_primarykey');
            const keys = Object.fromEntries(keyRows.map(k => [k.Z_NAME, k.Z_ENT]));
            
            console.log('📊 Analyzing notes structure...');
            
            // Get notes
            const notes = await database.all(`
                SELECT 
                    z_pk, zfolder, ztitle1 
                FROM ziccloudsyncingobject 
                WHERE 
                    z_ent = ? 
                    AND ztitle1 IS NOT NULL
            `, [keys.ICNote]);
            
            this.noteCount = notes.length;
            console.log(`📝 Found ${this.noteCount} notes to export`);
            
            // Process each note
            for (let note of notes) {
                try {
                    await this.processNote(database, keys, note);
                    this.parsedNotes++;
                    
                    if (this.parsedNotes % 10 === 0 || this.parsedNotes === this.noteCount) {
                        console.log(`⏳ Progress: ${this.parsedNotes}/${this.noteCount} notes processed`);
                    }
                } catch (error) {
                    console.warn(`⚠️  Failed to process note "${note.ZTITLE1}": ${error.message}`);
                }
            }
            
        } finally {
            database.close();
            // Clean up temporary database
            await fs.unlink(clonedDB).catch(() => {});
            await fs.unlink(clonedDB + '-shm').catch(() => {});
            await fs.unlink(clonedDB + '-wal').catch(() => {});
        }
        
        console.log(`✅ Export completed! ${this.parsedNotes} notes and ${this.attachmentCount} attachments exported to ${this.outputDir}`);
    }
    
    async processNote(database, keys, note) {
        // Get note data
        const row = await database.get(`
            SELECT 
                nd.z_pk, hex(nd.zdata) as zhexdata, zcso.ztitle1,
                zcreationdate1, zmodificationdate1
            FROM 
                zicnotedata AS nd,
                ziccloudsyncingobject AS zcso
            WHERE 
                zcso.z_pk = nd.znote
                AND zcso.z_pk = ?
        `, [note.Z_PK]);
        
        if (!row || !row.zhexdata) {
            console.warn(`⚠️  No data found for note: ${note.ZTITLE1}`);
            return;
        }
        
        // Extract text using protobuf decoder
        let content = '';
        try {
            const unzipped = zlib.gunzipSync(Buffer.from(row.zhexdata, 'hex'));
            content = await this.extractNoteContent(unzipped, note.ZTITLE1, database, keys, note.Z_PK);
        } catch (error) {
            console.warn(`⚠️  Failed to decode data for note "${note.ZTITLE1}": ${error.message}`);
            content = `# ${note.ZTITLE1}\n\n*Content could not be extracted*`;
        }
        
        // Create file
        const fileName = this.sanitizeFileName(note.ZTITLE1) + '.md';
        const filePath = path.join(this.outputDir, fileName);
        
        await fs.mkdir(path.dirname(filePath), { recursive: true });
        await fs.writeFile(filePath, content);
        
        // Set timestamps if available
        if (row.ZCREATIONDATE1 || row.ZMODIFICATIONDATE1) {
            const CORETIME_OFFSET = 978307200;
            const ctime = row.ZCREATIONDATE1 ? new Date((row.ZCREATIONDATE1 + CORETIME_OFFSET) * 1000) : new Date();
            const mtime = row.ZMODIFICATIONDATE1 ? new Date((row.ZMODIFICATIONDATE1 + CORETIME_OFFSET) * 1000) : ctime;
            
            await fs.utimes(filePath, ctime, mtime).catch(() => {});
        }
    }
    
    async extractNoteContent(buffer, title, database, keys, noteId) {
        try {
            // Try to decode as Document first
            const DocumentType = this.protobufRoot.lookupType('ciofecaforensics.Document');
            const document = DocumentType.decode(buffer);
            
            if (document.note && document.note.noteText) {
                return await this.formatNoteTextWithAttachments(document.note, title, database, keys, noteId);
            }
        } catch (documentError) {
            // If Document decoding fails, try MergableDataProto
            try {
                const MergableDataType = this.protobufRoot.lookupType('ciofecaforensics.MergableDataProto');
                const mergableData = MergableDataType.decode(buffer);
                
                // Look for note content in mergable data
                if (mergableData.mergableDataObject?.mergeableDataObjectData?.mergeableDataObjectEntry) {
                    const entries = mergableData.mergableDataObject.mergeableDataObjectData.mergeableDataObjectEntry;
                    
                    for (const entry of entries) {
                        if (entry.note && entry.note.noteText) {
                            return await this.formatNoteTextWithAttachments(entry.note, title, database, keys, noteId);
                        }
                        if (entry.orderedSet?.ordering?.array?.contents && entry.orderedSet.ordering.array.contents.noteText) {
                            return await this.formatNoteTextWithAttachments(entry.orderedSet.ordering.array.contents, title, database, keys, noteId);
                        }
                    }
                }
            } catch (mergableError) {
                console.warn(`⚠️  Protobuf decoding failed for "${title}": ${mergableError.message}`);
            }
        }
        
        // Fallback to basic text extraction
        return this.extractTextFallback(buffer, title);
    }
    
    async formatNoteTextWithAttachments(note, title, database, keys, noteId) {
        let content = `# ${title}\n\n`;
        let text = note.noteText || '';
        let attributeRuns = note.attributeRun || [];
        
        // Handle first line omission
        let skipFirstLine = this.omitFirstLine && text.includes('\n');
        let textOffset = 0;
        
        if (skipFirstLine) {
            const firstNewlineIndex = text.indexOf('\n');
            textOffset = firstNewlineIndex + 1;
        }
        
        // Process attribute runs to find attachments
        let processedText = '';
        let currentOffset = textOffset;
        
        for (const attr of attributeRuns) {
            // Add text before this attribute run
            const attrStart = currentOffset;
            const attrEnd = currentOffset + attr.length;
            
            if (attrStart < textOffset) {
                // Skip this run if it's in the first line we're omitting
                currentOffset = attrEnd;
                continue;
            }
            
            const textSegment = text.substring(attrStart, attrEnd);
            
            if (attr.attachmentInfo) {
                // Handle attachment
                const attachmentMarkdown = await this.processAttachment(attr.attachmentInfo, database, keys);
                processedText += attachmentMarkdown;
            } else {
                // Regular text
                processedText += textSegment;
            }
            
            currentOffset = attrEnd;
        }
        
        // Add any remaining text
        if (currentOffset < text.length) {
            processedText += text.substring(currentOffset);
        }
        
        // If no attribute runs or something went wrong, fall back to simple text
        if (attributeRuns.length === 0 || !processedText.trim()) {
            processedText = text.substring(textOffset);
        }
        
        // Clean up the text
        processedText = processedText
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .trim();
        
        content += processedText;
        return content;
    }
    
    formatNoteText(note, title) {
        let content = `# ${title}\n\n`;
        let text = note.noteText || '';
        
        // Handle first line omission
        if (this.omitFirstLine && text.includes('\n')) {
            const firstNewlineIndex = text.indexOf('\n');
            text = text.substring(firstNewlineIndex + 1);
        }
        
        // Clean up the text
        text = text
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .trim();
        
        content += text;
        return content;
    }
    
    extractTextFallback(buffer, title) {
        // Fallback: try to extract any readable text from the buffer
        let text = buffer.toString('utf8');
        
        // Remove control characters but keep newlines and tabs
        text = text.replace(/[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F-\x9F]/g, '').trim();
        
        // Look for readable text patterns (consecutive printable characters)
        const readableChunks = text.match(/[a-zA-Z0-9\s\.,!?;:'"()\-+=<>@#$%^&*{}[\]|\\\/]{10,}/g);
        
        if (!readableChunks || readableChunks.length === 0) {
            return `# ${title}\n\n*Note content could not be extracted - may contain rich content, attachments, or encrypted data*`;
        }
        
        let content = `# ${title}\n\n`;
        const extractedText = readableChunks.join('\n').trim();
        
        if (this.omitFirstLine && extractedText.includes('\n')) {
            const lines = extractedText.split('\n');
            if (lines.length > 1) {
                lines.shift();
                content += lines.join('\n');
            } else {
                content += extractedText;
            }
        } else {
            content += extractedText;
        }
        
        return content;
    }
    
    async processAttachment(attachmentInfo, database, keys) {
        const identifier = attachmentInfo.attachmentIdentifier;
        const typeUti = attachmentInfo.typeUti;
        
        try {
            switch (typeUti) {
                case this.ANAttachment.Hashtag:
                case this.ANAttachment.Mention:
                    const textRow = await database.get(`
                        SELECT zalttext FROM ziccloudsyncingobject 
                        WHERE zidentifier = ?
                    `, [identifier]);
                    return textRow?.ZALTTEXT || `#${identifier}`;
                    
                case this.ANAttachment.UrlCard:
                    const urlRow = await database.get(`
                        SELECT ztitle, zurlstring FROM ziccloudsyncingobject 
                        WHERE zidentifier = ?
                    `, [identifier]);
                    return urlRow ? `[**${urlRow.ZTITLE}**](${urlRow.ZURLSTRING})` : '[URL Card]';
                    
                case this.ANAttachment.Table:
                    return await this.processTable(identifier, database, keys);
                    
                case this.ANAttachment.Scan:
                case this.ANAttachment.ModifiedScan:
                case this.ANAttachment.Drawing:
                case this.ANAttachment.DrawingLegacy:
                case this.ANAttachment.DrawingLegacy2:
                    return await this.exportAttachmentFile(identifier, typeUti, database, keys);
                    
                default:
                    // Regular media file (image, video, audio, etc.)
                    const mediaRow = await database.get(`
                        SELECT zmedia FROM ziccloudsyncingobject 
                        WHERE zidentifier = ?
                    `, [identifier]);
                    if (mediaRow?.ZMEDIA) {
                        return await this.exportAttachmentFile(identifier, typeUti, database, keys, mediaRow.ZMEDIA);
                    }
                    return `\n\n*[Attachment: ${typeUti}]*\n\n`;
            }
        } catch (error) {
            console.warn(`⚠️  Failed to process attachment ${identifier}: ${error.message}`);
            return `\n\n*[Attachment processing failed: ${typeUti}]*\n\n`;
        }
    }
    
    async processTable(identifier, database, keys) {
        try {
            const row = await database.get(`
                SELECT hex(zmergeabledata1) as zhexdata 
                FROM ziccloudsyncingobject 
                WHERE zidentifier = ?
            `, [identifier]);
            
            if (!row || !row.zhexdata) {
                return '\n\n*[Table data not found]*\n\n';
            }
            
            // Decode the protobuf table data
            const unzipped = zlib.gunzipSync(Buffer.from(row.zhexdata, 'hex'));
            const tableMarkdown = await this.convertTableData(unzipped);
            return tableMarkdown || '\n\n*[Table could not be processed]*\n\n';
            
        } catch (error) {
            console.warn(`⚠️  Failed to process table ${identifier}: ${error.message}`);
            return '\n\n*[Table processing failed]*\n\n';
        }
    }
    
    async convertTableData(buffer) {
        try {
            // Try to decode as MergableDataProto
            const MergableDataType = this.protobufRoot.lookupType('ciofecaforensics.MergableDataProto');
            const mergableData = MergableDataType.decode(buffer);
            
            if (!mergableData.mergableDataObject?.mergeableDataObjectData) {
                return null;
            }
            
            const data = mergableData.mergableDataObject.mergeableDataObjectData;
            const keys = data.mergeableDataObjectKeyItem || [];
            const types = data.mergeableDataObjectTypeItem || [];
            const uuids = (data.mergeableDataObjectUuidItem || []).map(uuid => Buffer.from(uuid).toString('hex'));
            const objects = data.mergeableDataObjectEntry || [];
            
            // Find the root table object
            const root = objects.find(e => e.customMap && types[e.customMap.type] === 'com.apple.notes.ICTable');
            if (!root) return null;
            
            let rowLocations = {};
            let rowCount = 0;
            let columnLocations = {};
            let columnCount = 0;
            let cellData = null;
            
            // Parse the table structure
            for (const entry of root.customMap.mapEntry) {
                const object = objects[entry.value.objectIndex];
                const keyName = keys[entry.key];
                
                switch (keyName) {
                    case 'crRows':
                        [rowLocations, rowCount] = this.findLocations(object, objects, uuids);
                        break;
                    case 'crColumns':
                        [columnLocations, columnCount] = this.findLocations(object, objects, uuids);
                        break;
                    case 'cellColumns':
                        cellData = object;
                        break;
                }
            }
            
            if (!cellData) return null;
            
            // Build the table data
            const table = Array(rowCount).fill(0).map(() => Array(columnCount));
            
            for (const column of cellData.dictionary.element) {
                const columnLocation = columnLocations[this.getTargetUuid(column.key, objects, uuids)];
                const rowData = objects[column.value.objectIndex];
                
                if (rowData?.dictionary?.element) {
                    for (const row of rowData.dictionary.element) {
                        const rowLocation = rowLocations[this.getTargetUuid(row.key, objects, uuids)];
                        const rowContent = objects[row.value.objectIndex];
                        
                        if (rowLocation < rowCount && columnLocation < columnCount && rowContent) {
                            // Extract text from the cell content
                            let cellText = '';
                            if (rowContent.note && rowContent.note.noteText) {
                                cellText = rowContent.note.noteText.trim();
                            }
                            table[rowLocation][columnLocation] = cellText || '';
                        }
                    }
                }
            }
            
            // Convert to markdown table
            let markdown = '\n';
            for (let i = 0; i < table.length; i++) {
                const row = table[i].map(cell => (cell || '').replace(/\|/g, '\\|')); // Escape pipes
                markdown += `| ${row.join(' | ')} |\n`;
                if (i === 0) {
                    markdown += `|${' -- |'.repeat(table[0].length)}\n`;
                }
            }
            return markdown + '\n';
            
        } catch (error) {
            console.warn(`⚠️  Table protobuf decoding failed: ${error.message}`);
            return null;
        }
    }
    
    findLocations(object, objects, uuids) {
        let ordering = [];
        let indices = {};
        
        if (object.orderedSet?.ordering?.array?.attachment) {
            for (const element of object.orderedSet.ordering.array.attachment) {
                ordering.push(Buffer.from(element.uuid).toString('hex'));
            }
        }
        
        if (object.orderedSet?.ordering?.contents?.element) {
            for (const element of object.orderedSet.ordering.contents.element) {
                const key = this.getTargetUuid(element.key, objects, uuids);
                const value = this.getTargetUuid(element.value, objects, uuids);
                
                indices[value] = ordering.indexOf(key);
            }
        }
        
        return [indices, ordering.length];
    }
    
    getTargetUuid(entry, objects, uuids) {
        const reference = objects[entry.objectIndex];
        if (reference?.customMap?.mapEntry?.[0]?.value?.unsignedIntegerValue !== undefined) {
            const uuidIndex = reference.customMap.mapEntry[0].value.unsignedIntegerValue;
            return uuids[uuidIndex];
        }
        return '';
    }
    
    async exportAttachmentFile(identifier, typeUti, database, keys, mediaId = null) {
        try {
            let row, sourcePath, outName, outExt;
            
            if (typeUti === this.ANAttachment.ModifiedScan) {
                row = await database.get(`
                    SELECT zidentifier, zfallbackpdfgeneration, zcreationdate, zmodificationdate 
                    FROM ziccloudsyncingobject 
                    WHERE z_ent = ? AND zidentifier = ?
                `, [keys.ICAttachment, identifier]);
                
                sourcePath = path.join('FallbackPDFs', row.ZIDENTIFIER, row.ZFALLBACKPDFGENERATION || '', 'FallbackPDF.pdf');
                outName = 'Scan';
                outExt = 'pdf';
                
            } else if (typeUti === this.ANAttachment.Scan) {
                row = await database.get(`
                    SELECT zidentifier, zsizeheight, zsizewidth, zcreationdate, zmodificationdate 
                    FROM ziccloudsyncingobject 
                    WHERE z_ent = ? AND zidentifier = ?
                `, [keys.ICAttachment, identifier]);
                
                sourcePath = path.join('Previews', `${row.ZIDENTIFIER}-1-${row.ZSIZEWIDTH}x${row.ZSIZEHEIGHT}-0.jpeg`);
                outName = 'Scan Page';
                outExt = 'jpg';
                
            } else if ([this.ANAttachment.Drawing, this.ANAttachment.DrawingLegacy, this.ANAttachment.DrawingLegacy2].includes(typeUti)) {
                row = await database.get(`
                    SELECT zidentifier, zfallbackimagegeneration, zcreationdate, zmodificationdate, zhandwritingsummary 
                    FROM ziccloudsyncingobject 
                    WHERE z_ent = ? AND zidentifier = ?
                `, [keys.ICAttachment, identifier]);
                
                if (row.ZFALLBACKIMAGEGENERATION) {
                    sourcePath = path.join('FallbackImages', row.ZIDENTIFIER, row.ZFALLBACKIMAGEGENERATION, 'FallbackImage.png');
                } else {
                    sourcePath = path.join('FallbackImages', `${row.ZIDENTIFIER}.jpg`);
                }
                outName = 'Drawing';
                outExt = 'png';
                
            } else if (mediaId) {
                // Regular media file
                row = await database.get(`
                    SELECT a.zidentifier, a.zfilename, a.zgeneration1, b.zcreationdate, b.zmodificationdate 
                    FROM ziccloudsyncingobject AS a, ziccloudsyncingobject AS b 
                    WHERE a.z_ent = ? AND a.z_pk = ? AND a.z_pk = b.zmedia
                `, [keys.ICMedia, mediaId]);
                
                if (!row) return `\n\n*[Media file not found]*\n\n`;
                
                sourcePath = path.join('Media', row.ZIDENTIFIER, row.ZGENERATION1 || '', row.ZFILENAME);
                const parts = row.ZFILENAME.split('.');
                outExt = parts.length > 1 ? parts.pop() : 'bin';
                outName = parts.join('.') || 'attachment';
            }
            
            if (!row || !sourcePath) {
                return `\n\n*[Attachment data not found]*\n\n`;
            }
            
            // Try to copy the attachment file from any available account
            let fullSourcePath = null;
            
            // Try each account directory
            for (const account of this.accounts) {
                if (account.mediaPath) {
                    const testPath = path.join(account.path, sourcePath);
                    const exists = await fs.access(testPath, fs.constants.F_OK).then(() => true).catch(() => false);
                    if (exists) {
                        fullSourcePath = testPath;
                        break;
                    }
                }
            }
            
            // Fallback to old method
            if (!fullSourcePath) {
                fullSourcePath = path.join(this.dataPath, sourcePath);
            }
            
            const attachmentsDir = path.join(this.outputDir, 'attachments');
            await fs.mkdir(attachmentsDir, { recursive: true });
            
            // Generate unique filename
            let finalName = `${this.sanitizeFileName(outName)}.${outExt}`;
            let counter = 1;
            let finalPath = path.join(attachmentsDir, finalName);
            
            while (await fs.access(finalPath, fs.constants.F_OK).then(() => true).catch(() => false)) {
                finalName = `${this.sanitizeFileName(outName)}_${counter}.${outExt}`;
                finalPath = path.join(attachmentsDir, finalName);
                counter++;
            }
            
            try {
                await fs.copyFile(fullSourcePath, finalPath);
                this.attachmentCount++;
                console.log(`📎 Exported attachment: ${finalName}`);
                
                // Set timestamps if available
                if (row.ZCREATIONDATE || row.ZMODIFICATIONDATE) {
                    const CORETIME_OFFSET = 978307200;
                    const ctime = row.ZCREATIONDATE ? new Date((row.ZCREATIONDATE + CORETIME_OFFSET) * 1000) : new Date();
                    const mtime = row.ZMODIFICATIONDATE ? new Date((row.ZMODIFICATIONDATE + CORETIME_OFFSET) * 1000) : ctime;
                    await fs.utimes(finalPath, ctime, mtime).catch(() => {});
                }
                
                // Return markdown link with handwriting summary if available
                let link = `\n\n![${outName}](attachments/${finalName})\n\n`;
                
                if (this.includeHandwriting && row.ZHANDWRITINGSUMMARY) {
                    link = `\n\n> [!note] Handwriting\n> ${row.ZHANDWRITINGSUMMARY.replace(/\n/g, '\n> ')}\n\n![${outName}](attachments/${finalName})\n\n`;
                }
                
                return link;
                
            } catch (copyError) {
                console.warn(`⚠️  Could not copy attachment file ${sourcePath}: ${copyError.message}`);
                return `\n\n*[Attachment file could not be exported: ${finalName}]*\n\n`;
            }
            
        } catch (error) {
            console.warn(`⚠️  Failed to export attachment file: ${error.message}`);
            return `\n\n*[Attachment export failed]*\n\n`;
        }
    }
    
    sanitizeFileName(name) {
        return name.replace(/[<>:"/\\|?*]/g, '').replace(/\s+/g, ' ').trim();
    }
}

async function main() {
    const outputDir = process.argv[2] || './notes';
    const omitFirstLine = process.argv[3] !== 'false';
    const importTrashed = process.argv[4] === 'true';
    const includeHandwriting = process.argv[5] === 'true';
    
    console.log('🚀 Starting Apple Notes export...');
    console.log(`📁 Output directory: ${outputDir}`);
    console.log(`⚙️  Options: omitFirstLine=${omitFirstLine}, importTrashed=${importTrashed}, includeHandwriting=${includeHandwriting}`);
    console.log('');
    
    try {
        const exporter = new AppleNotesExporter(outputDir, {
            omitFirstLine,
            importTrashed,
            includeHandwriting
        });
        
        await exporter.export();
    } catch (error) {
        console.error('❌ Export failed:', error.message);
        process.exit(1);
    }
}

main();
EOF
    
    echo -e "${BLUE}Running automated Apple Notes export...${NC}"
    cd "$SCRIPT_DIR"
    node standalone-export.js "$output_dir" "$OMIT_FIRST_LINE" "$IMPORT_TRASHED" "$INCLUDE_HANDWRITING"
    
    if [ $? -eq 0 ]; then
        echo ""
        echo -e "${GREEN}✅ Export completed successfully!${NC}"
        echo -e "${GREEN}📁 Your Apple Notes have been exported to: $output_dir${NC}"
    else
        echo ""
        echo -e "${RED}❌ Export failed. Please check the output above for details.${NC}"
        exit 1
    fi
}

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        -o|--output)
            OUTPUT_DIR="$2"
            shift 2
            ;;
        --include-first-line)
            OMIT_FIRST_LINE=false
            shift
            ;;
        --include-trashed)
            IMPORT_TRASHED=true
            shift
            ;;
        --include-handwriting)
            INCLUDE_HANDWRITING=true
            shift
            ;;
        -h|--help)
            print_help
            exit 0
            ;;
        *)
            echo -e "${RED}Unknown option: $1${NC}"
            print_help
            exit 1
            ;;
    esac
done

# Validate required arguments
if [[ -z "$OUTPUT_DIR" ]]; then
    echo -e "${RED}Error: Output directory is required${NC}"
    echo ""
    print_help
    exit 1
fi

# Convert to absolute path
OUTPUT_DIR="$(cd "$(dirname "$OUTPUT_DIR")" && pwd)/$(basename "$OUTPUT_DIR")"

echo -e "${GREEN}Apple Notes to Markdown Export Script${NC}"
echo "=================================="
echo "Output directory: $OUTPUT_DIR"
echo "Omit first line: $OMIT_FIRST_LINE"
echo "Import trashed notes: $IMPORT_TRASHED"
echo "Include handwriting: $INCLUDE_HANDWRITING"
echo ""

# Run the export process
check_requirements
install_dependencies
build_project
export_notes "$OUTPUT_DIR"

echo ""
echo -e "${GREEN}Setup complete!${NC}"
echo "Follow the manual steps above to complete the export."