function getOfficialSpreadsheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('ERP_SPREADSHEET_ID');
  if (!id || id.indexOf('여기에_') === 0) {
    throw new Error('스크립트 속성에 ERP_SPREADSHEET_ID (공식 전화번호부 Google Sheet ID)를 설정해야 합니다.');
  }
  return id;
}
const OFFICIAL_SHEET_NAME = '비상연락망';
const DEFAULT_LABEL_NAME = '회사 전화번호부';
const USER_SETTINGS_KEY = 'contactSyncWebAppSettings.v1';
const HEADER_SCAN_ROWS = 20;
// PREVIEW_LIMIT removed — full preview is returned
const CONTACT_GROUP_MEMBER_CHUNK_SIZE = 100;

const HEADER_ALIASES = {
    name: ['이름', '성명', '직원명', '담당자', 'name', 'full name', 'fullname'],
    phone: ['전화번호', '전화', '휴대폰', '핸드폰', '연락처', '번호', 'phone', 'mobile', 'cell'],
    email: ['이메일', '메일', '메일주소', 'email', 'e-mail', 'mail'],
    company: ['회사', '회사명', '소속회사', '법인', 'company', 'organization', 'org'],
    department: ['부서', '팀', '소속', 'department', 'team', 'division'],
    title: ['직급', '직책', '직위', '포지션', 'title', 'position', 'role'],
    note: ['메모', '비고', '참고', 'note', 'memo', 'remarks'],
};

function doGet() {
    return HtmlService
        .createHtmlOutputFromFile('WebApp')
        .setTitle('회사 전화번호부 동기화')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialState() {
    const analysis = analyzeOfficialSheet_();
    const settings = getUserSettings_();
    const mapping = Object.assign({}, analysis.detectedMapping, settings.mapping || {});
    const template = settings.template || buildDefaultTemplate_(analysis.headers);
    const options = Object.assign({
        replaceLabelContacts: true,
        skipRowsWithoutPhone: true,
        fillMergedLikeBlanks: true,
        saveSettings: true,
        showLowestDepartmentOnly: false,
        smartSync: true,
        keepMissingContacts: false,
    }, settings.options || {});

    return {
        spreadsheetName: analysis.spreadsheetName,
        sheetName: analysis.sheetName,
        lastUpdated: analysis.lastUpdated,
        headerRow: analysis.headerRow,
        startColumn: analysis.startColumn,
        endColumn: analysis.endColumn,
        totalRows: analysis.rows.length,
        headers: analysis.headers,
        mapping,
        template,
        labelName: settings.labelName || DEFAULT_LABEL_NAME,
        options,
        preview: buildPreview_(analysis, mapping, template, options, {}),
        warnings: analysis.warnings,
        autoSyncEnabled: checkAutoSyncStatus_(),
    };
}

function getPreview(payload) {
    const analysis = analyzeOfficialSheet_();
    return buildPreview_(analysis, payload.mapping || {}, payload.template || '', payload.options || {}, payload.rowOverrides || {}, payload.addedRows || [], payload.deletedRows || []);
}

function saveUserSettings(payload) {
    const settings = {
        mapping: payload.mapping || {},
        template: payload.template || '{이름}',
        labelName: payload.labelName || DEFAULT_LABEL_NAME,
        options: payload.options || {},
        updatedAt: new Date().toISOString(),
    };
    PropertiesService.getUserProperties().setProperty(USER_SETTINGS_KEY, JSON.stringify(settings));
    return { ok: true };
}

function checkAutoSyncStatus_() {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'syncMyContactsAuto') {
            return true;
        }
    }
    return false;
}

function toggleAutoSync(payload) {
    const enable = !!payload.enable;
    
    // 저장된 설정 업데이트 (자동 동기화 시 사용할 설정)
    if (enable) {
        saveUserSettings({
            mapping: payload.mapping || {},
            template: payload.template || '',
            labelName: payload.labelName || DEFAULT_LABEL_NAME,
            options: payload.options || {}
        });
    }
    
    // 중복 방지를 위해 기존 트리거 모두 삭제
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'syncMyContactsAuto') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    
    // 활성화인 경우 매일 새벽 2시경 동작하는 트리거 생성
    if (enable) {
        ScriptApp.newTrigger('syncMyContactsAuto')
            .timeBased()
            .everyDays(1)
            .atHour(2) // 새벽 2시 ~ 3시 사이 무작위로 실행됨
            .create();
    }
    
    return { enabled: enable };
}

function syncMyContactsAuto() {
    const settings = getUserSettings_();
    if (!settings || !settings.mapping || Object.keys(settings.mapping).length === 0) {
        console.warn('저장된 설정이 없어 자동 동기화를 건너뜁니다.');
        return;
    }
    
    try {
        const results = syncMyContacts({
            mapping: settings.mapping,
            template: settings.template,
            labelName: settings.labelName,
            options: settings.options,
            rowOverrides: {},
            addedRows: [],
            deletedRows: []
        });
        console.log(`자동 동기화 완료: 추가 ${results.created}건, 수정 ${results.updated}건, 삭제 ${results.deleted}건, 제외 ${results.skipped}건, 에러 ${results.errors.length}건`);
    } catch(e) {
        console.error('자동 동기화 중 오류 발생: ' + e.message);
    }
}

function syncMyContacts(payload) {
    const analysis = analyzeOfficialSheet_();
    const mapping = payload.mapping || {};
    const template = payload.template || '{이름}';
    const labelName = String(payload.labelName || DEFAULT_LABEL_NAME).trim();
    const options = Object.assign({
        replaceLabelContacts: true,
        skipRowsWithoutPhone: true,
        fillMergedLikeBlanks: true,
        saveSettings: true,
        showLowestDepartmentOnly: false,
        smartSync: true,
        keepMissingContacts: false,
    }, payload.options || {});

    const rowOverrides = payload.rowOverrides || {};
    const addedRows = payload.addedRows || [];
    const deletedRows = payload.deletedRows || [];

    validateRequiredMapping_(mapping);
    if (!labelName) {
        throw new Error('라벨 이름을 입력해야 합니다.');
    }

    if (options.saveSettings) {
        saveUserSettings({ mapping, template, labelName, options });
    }

    const contactGroup = findOrCreateContactGroup_(labelName);
    
    // 스마트 동기화를 위해 기존 연락처 가져오기
    const existingContacts = fetchExistingContacts_(contactGroup.resourceName);
    
    let rows = options.fillMergedLikeBlanks
        ? fillMergedLikeBlankRows_(analysis.rows, analysis.headers, mapping, options)
        : analysis.rows;

    // Apply overrides and deletions
    rows = rows.map((row, index) => {
        const rowNumber = analysis.headerRow + 1 + index;
        if (rowOverrides[rowNumber]) {
            return analysis.headers.map(h => rowOverrides[rowNumber][h] != null ? rowOverrides[rowNumber][h] : row[analysis.headers.indexOf(h)]);
        }
        return row;
    }).filter((row, index) => {
        const rowNumber = analysis.headerRow + 1 + index;
        return !deletedRows.includes(rowNumber);
    });

    // Append added rows
    addedRows.forEach(addedRecord => {
        const row = analysis.headers.map(h => addedRecord[h] || '');
        rows.push(row);
    });

    const results = {
        created: 0,
        updated: 0,
        deleted: 0,
        skipped: 0,
        errors: [],
        synced: [],
    };

    const targetPeople = [];
    rows.forEach((row, index) => {
        const rowNumber = analysis.headerRow + 1 + index;
        const record = rowToRecord_(row, analysis.headers);
        const phone = normalizePhone_(getMappedValue_(record, mapping.phone));
        const name = String(getMappedValue_(record, mapping.name) || '').trim();

        if (!name || !phone) {
            results.skipped += 1;
            if (!name || (!phone && !options.skipRowsWithoutPhone)) {
                results.errors.push({ rowNumber, message: '필수값(이름/전화번호)이 부족합니다.' });
            }
            return;
        }

        const displayName = renderTemplate_(template, record).trim() || name;
        const person = buildPerson_(record, mapping, displayName);
        targetPeople.push({ rowNumber, displayName, phone, person });
    });

    // Diffing 로직
    const toCreate = [];
    const toUpdate = {}; // { resourceName: Person }
    const matchedExistingResourceNames = new Set();
    const updateEtagMap = {}; // for storing etags

    // 매칭을 위해 기존 연락처 맵 생성 (전화번호 기준, 없으면 이름 기준)
    const existingByPhone = {};
    const existingByName = {};
    
    existingContacts.forEach(person => {
        const phoneObj = (person.phoneNumbers || []).find(p => p.value);
        if (phoneObj) {
            existingByPhone[normalizeDiffPhone_(phoneObj.value)] = person;
        }
        const nameObj = (person.names || []).find(n => n.givenName || n.displayName);
        if (nameObj) {
            existingByName[nameObj.givenName || nameObj.displayName] = person;
        }
    });

    targetPeople.forEach(target => {
        const diffPhone = normalizeDiffPhone_(target.phone);
        const diffName = (target.person.names[0].givenName || target.person.names[0].displayName);
        
        let existingMatch = existingByPhone[diffPhone];
        
        if (existingMatch) {
            // 매칭된 항목이 변경되었는지 확인
            const isChanged = checkPersonChanged_(existingMatch, target.person);
            if (isChanged) {
                target.person.etag = existingMatch.etag;
                toUpdate[existingMatch.resourceName] = target.person;
                updateEtagMap[existingMatch.resourceName] = existingMatch.etag;
            } else {
                results.skipped += 1;
            }
            matchedExistingResourceNames.add(existingMatch.resourceName);
            results.synced.push({ rowNumber: target.rowNumber, displayName: target.displayName, phone: target.phone });
        } else {
            // 새로 생성
            toCreate.push({ contactPerson: target.person });
            results.synced.push({ rowNumber: target.rowNumber, displayName: target.displayName, phone: target.phone });
        }
    });

    // 삭제할 연락처 찾기
    const toDelete = [];
    if (!options.keepMissingContacts) {
        existingContacts.forEach(person => {
            if (!matchedExistingResourceNames.has(person.resourceName)) {
                toDelete.push(person.resourceName);
            }
        });
    }

    // Batch API 호출
    const updateMask = 'names,phoneNumbers,emailAddresses,organizations,biographies';
    
    // 1. Delete
    if (toDelete.length > 0) {
        for (let i = 0; i < toDelete.length; i += 500) {
            try {
                People.People.batchDeleteContacts({ resourceNames: toDelete.slice(i, i + 500) });
                results.deleted += toDelete.slice(i, i + 500).length;
            } catch (e) {
                results.errors.push({ rowNumber: '-', message: '일괄 삭제 실패: ' + e.message });
            }
        }
    }

    // 2. Create
    if (toCreate.length > 0) {
        const createdResourceNames = [];
        for (let i = 0; i < toCreate.length; i += 200) {
            try {
                const chunk = toCreate.slice(i, i + 200);
                const response = People.People.batchCreateContacts({ contacts: chunk, readMask: updateMask });
                if (response.createdPeople) {
                    response.createdPeople.forEach(cp => {
                        if (cp.person && cp.person.resourceName) {
                            createdResourceNames.push(cp.person.resourceName);
                            results.created += 1;
                        }
                    });
                }
            } catch (e) {
                results.errors.push({ rowNumber: '-', message: '일괄 생성 실패: ' + e.message });
            }
        }
        if (createdResourceNames.length > 0) {
            addContactsToGroup_(contactGroup.resourceName, createdResourceNames);
        }
    }

    // 3. Update
    const updateKeys = Object.keys(toUpdate);
    if (updateKeys.length > 0) {
        for (let i = 0; i < updateKeys.length; i += 200) {
            try {
                const chunkKeys = updateKeys.slice(i, i + 200);
                const contactsObj = {};
                chunkKeys.forEach(k => { contactsObj[k] = toUpdate[k]; });
                
                const response = People.People.batchUpdateContacts({
                    contacts: contactsObj,
                    updateMask: updateMask,
                    readMask: updateMask
                });
                
                if (response.updateResult) {
                    results.updated += Object.keys(response.updateResult).length;
                }
            } catch (e) {
                results.errors.push({ rowNumber: '-', message: '일괄 업데이트 실패: ' + e.message });
            }
        }
    }

    return results;
}

function fetchExistingContacts_(contactGroupResourceName) {
    const group = People.ContactGroups.get(contactGroupResourceName, { maxMembers: 10000 });
    const resourceNames = group.memberResourceNames || [];
    
    const existingContacts = [];
    for (let i = 0; i < resourceNames.length; i += 200) {
        const chunk = resourceNames.slice(i, i + 200);
        const response = People.People.getBatchGet({
            resourceNames: chunk,
            personFields: 'names,phoneNumbers,emailAddresses,organizations,biographies'
        });
        if (response.responses) {
            response.responses.forEach(res => {
                if (res.person) existingContacts.push(res.person);
            });
        }
    }
    return existingContacts;
}

function normalizeDiffPhone_(phone) {
    return String(phone || '').replace(/[^\d]/g, '');
}

function checkPersonChanged_(existing, target) {
    // Check Name
    const extName = (existing.names && existing.names[0]) ? (existing.names[0].givenName || existing.names[0].displayName || '') : '';
    const tgtName = (target.names && target.names[0]) ? (target.names[0].givenName || target.names[0].displayName || '') : '';
    if (extName !== tgtName) return true;

    // Check Phone
    const extPhone = (existing.phoneNumbers && existing.phoneNumbers[0]) ? normalizeDiffPhone_(existing.phoneNumbers[0].value) : '';
    const tgtPhone = (target.phoneNumbers && target.phoneNumbers[0]) ? normalizeDiffPhone_(target.phoneNumbers[0].value) : '';
    if (extPhone !== tgtPhone) return true;

    // Check Email
    const extEmail = (existing.emailAddresses && existing.emailAddresses[0]) ? (existing.emailAddresses[0].value || '') : '';
    const tgtEmail = (target.emailAddresses && target.emailAddresses[0]) ? (target.emailAddresses[0].value || '') : '';
    if (extEmail !== tgtEmail) return true;

    // Check Organization
    const extOrg = (existing.organizations && existing.organizations[0]) ? existing.organizations[0] : {};
    const tgtOrg = (target.organizations && target.organizations[0]) ? target.organizations[0] : {};
    if ((extOrg.name || '') !== (tgtOrg.name || '')) return true;
    if ((extOrg.department || '') !== (tgtOrg.department || '')) return true;
    if ((extOrg.title || '') !== (tgtOrg.title || '')) return true;

    // Check Bio
    const extBio = (existing.biographies && existing.biographies[0]) ? (existing.biographies[0].value || '') : '';
    const tgtBio = (target.biographies && target.biographies[0]) ? (target.biographies[0].value || '') : '';
    if (extBio !== tgtBio) return true;

    return false;
}

function analyzeOfficialSheet_() {
    const spreadsheetId = getOfficialSpreadsheetId_();
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = null;
    
    if (OFFICIAL_SHEET_NAME) {
        sheet = spreadsheet.getSheetByName(OFFICIAL_SHEET_NAME);
    }
    
    // 지정된 이름의 시트가 없으면 첫 번째 시트로 대체
    if (!sheet) {
        sheet = spreadsheet.getSheets()[0];
    }

    if (!sheet) {
        throw new Error('공식 전화번호부 시트를 찾지 못했습니다.');
    }

    const values = sheet.getDataRange().getDisplayValues();
    if (!values.length) {
        throw new Error('공식 전화번호부 시트에 데이터가 없습니다.');
    }

    const headerInfo = detectHeaderRow_(values);
    const headers = headerInfo.headers;
    const rows = values
        .slice(headerInfo.rowIndex + 1)
        .map(row => row.slice(headerInfo.startCol, headerInfo.endCol + 1))
        .filter(row => row.some(cell => String(cell).trim() !== ''));

    return {
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastUpdated: PropertiesService.getScriptProperties().getProperty('LAST_SYNC_TIME') || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm'),
        headerRow: headerInfo.rowIndex + 1,
        startColumn: headerInfo.startCol + 1,
        endColumn: headerInfo.endCol + 1,
        headers,
        detectedMapping: detectMapping_(headers),
        rows,
        warnings: headerInfo.warnings,
    };
}

function detectHeaderRow_(values) {
    let best = { rowIndex: 0, score: -1, headers: values[0] || [], aliasIndexes: [] };
    const maxRows = Math.min(values.length, HEADER_SCAN_ROWS);

    for (let i = 0; i < maxRows; i += 1) {
        const row = values[i].map(value => String(value).trim());
        const nonEmptyCount = row.filter(Boolean).length;
        const aliasIndexes = [];
        const aliasScore = row.reduce((score, cell, index) => {
            const cellScore = scoreHeaderCell_(cell);
            if (cellScore) aliasIndexes.push(index);
            return score + cellScore;
        }, 0);
        const score = aliasScore * 10 + nonEmptyCount;
        if (score > best.score) {
            best = { rowIndex: i, score, headers: row, aliasIndexes };
        }
    }

    const bounds = detectTableColumnBounds_(values, best.rowIndex, best.headers, best.aliasIndexes);
    const headers = makeUniqueHeaders_(best.headers.slice(bounds.startCol, bounds.endCol + 1));
    const warnings = [];
    if (best.score < 20) {
        warnings.push('헤더를 확실하게 찾지 못했습니다. 열 매핑을 직접 확인해주세요.');
    }

    return {
        rowIndex: best.rowIndex,
        startCol: bounds.startCol,
        endCol: bounds.endCol,
        headers,
        warnings,
    };
}

function detectTableColumnBounds_(values, headerRowIndex, headerRow, aliasIndexes) {
    const nonEmptyHeaderIndexes = headerRow
        .map((cell, index) => String(cell).trim() ? index : -1)
        .filter(index => index >= 0);
    const anchors = aliasIndexes.length ? aliasIndexes : nonEmptyHeaderIndexes;

    if (!anchors.length) {
        return { startCol: 0, endCol: headerRow.length - 1 };
    }

    let startCol = Math.min.apply(null, anchors);
    let endCol = Math.max.apply(null, anchors);

    while (startCol > 0 && columnLooksRelated_(values, headerRowIndex, startCol - 1)) {
        startCol -= 1;
    }
    while (endCol < headerRow.length - 1 && columnLooksRelated_(values, headerRowIndex, endCol + 1)) {
        endCol += 1;
    }

    while (startCol <= endCol && !String(headerRow[startCol] || '').trim()) startCol += 1;
    while (endCol >= startCol && !String(headerRow[endCol] || '').trim()) endCol -= 1;

    return { startCol, endCol };
}

function columnLooksRelated_(values, headerRowIndex, columnIndex) {
    const header = String(values[headerRowIndex][columnIndex] || '').trim();
    if (header) return true;

    const sampleRows = values.slice(headerRowIndex + 1, Math.min(values.length, headerRowIndex + 12));
    const nonEmptyCount = sampleRows.filter(row => String(row[columnIndex] || '').trim()).length;
    return nonEmptyCount >= Math.min(3, sampleRows.length);
}

function scoreHeaderCell_(cell) {
    const normalized = normalizeHeader_(cell);
    if (!normalized) return 0;
    return Object.keys(HEADER_ALIASES).some(key =>
        HEADER_ALIASES[key].some(alias => normalizeHeader_(alias) === normalized)
    ) ? 1 : 0;
}

function detectMapping_(headers) {
    const mapping = {};
    headers.forEach(header => {
        const normalized = normalizeHeader_(header);
        Object.keys(HEADER_ALIASES).forEach(field => {
            if (mapping[field]) return;
            const matched = HEADER_ALIASES[field].some(alias => normalizeHeader_(alias) === normalized);
            if (matched) mapping[field] = header;
        });
    });
    return mapping;
}

function buildPreview_(analysis, mapping, template, options, rowOverrides, addedRows, deletedRows) {
    rowOverrides = rowOverrides || {};
    addedRows = addedRows || [];
    deletedRows = deletedRows || [];

    let previewRows = options && options.fillMergedLikeBlanks
        ? fillMergedLikeBlankRows_(analysis.rows, analysis.headers, mapping, options)
        : analysis.rows;

    const result = [];

    previewRows.forEach((row, index) => {
        const rowNumber = analysis.headerRow + 1 + index;
        if (deletedRows.includes(rowNumber)) return;

        let effectiveRow = row;
        let overridden = false;
        if (rowOverrides[rowNumber]) {
            effectiveRow = analysis.headers.map(h => rowOverrides[rowNumber][h] != null ? rowOverrides[rowNumber][h] : row[analysis.headers.indexOf(h)]);
            overridden = true;
        }

        const record = rowToRecord_(effectiveRow, analysis.headers);
        const displayName = renderTemplate_(template, record).trim();
        result.push({
            rowNumber,
            displayName,
            record,
            phone: getMappedValue_(record, mapping.phone),
            email: getMappedValue_(record, mapping.email),
            valid: Boolean(displayName && getMappedValue_(record, mapping.phone)),
            overridden,
            added: false,
        });
    });

    // Append added rows at the end
    addedRows.forEach((addedRecord, i) => {
        const displayName = renderTemplate_(template, addedRecord).trim();
        result.push({
            rowNumber: `+${i + 1}`,
            displayName,
            record: addedRecord,
            phone: getMappedValue_(addedRecord, mapping.phone),
            email: getMappedValue_(addedRecord, mapping.email),
            valid: Boolean(displayName && getMappedValue_(addedRecord, mapping.phone)),
            overridden: false,
            added: true,
        });
    });

    return result;
}

function fillMergedLikeBlankRows_(rows, headers, mapping, options = {}) {
    const companyIdx = mapping.company ? headers.indexOf(mapping.company) : -1;
    const deptIdx = mapping.department ? headers.indexOf(mapping.department) : -1;
    const titleIdx = mapping.title ? headers.indexOf(mapping.title) : -1;

    // 부서(dept) 컬럼부터 직책(title) 컬럼 이전까지의 모든 컬럼을 부서 계층 컬럼으로 인식합니다.
    const deptCols = [];
    if (deptIdx >= 0) {
        deptCols.push(deptIdx);
        if (titleIdx > deptIdx) {
            for (let i = deptIdx + 1; i < titleIdx; i++) {
                deptCols.push(i);
            }
        }
    }

    const fillableIndexes = [];
    if (companyIdx >= 0) fillableIndexes.push(companyIdx);
    deptCols.forEach(idx => fillableIndexes.push(idx));

    const uniqueFillable = fillableIndexes.filter((idx, i, all) => all.indexOf(idx) === i);
    const carry = {};

    return rows.map(row => {
        const nextRow = row.slice();

        // 1. 대표이사, 고문 등: 부서 컬럼에 기재되어 있고 직책이 비어있는 경우 직책 컬럼으로 이동
        if (deptIdx >= 0 && titleIdx >= 0) {
            const rawDept = String(nextRow[deptIdx] || '').trim();
            const rawTitle = String(nextRow[titleIdx] || '').trim();
            const isExecutiveTitle = /대표이사|고문|부사장|전무|상무|감사/.test(rawDept);
            
            if (rawDept && !rawTitle && isExecutiveTitle) {
                nextRow[titleIdx] = rawDept;
                nextRow[deptIdx] = ''; 
            }
        }

        // 2. 병합된 셀(빈 셀)에 대해 상단 값 채우기 (Carry Down)
        // 회사 -> 부서계층 간의 종속성을 고려하여, 상위 계층에 명시적인 값이 있으면 하위 계층의 carry를 초기화합니다.
        const companyValue = companyIdx >= 0 ? String(nextRow[companyIdx] || '').trim() : '';
        if (companyValue) {
            deptCols.forEach(idx => delete carry[idx]);
        }

        for (let i = 0; i < deptCols.length; i++) {
            const idx = deptCols[i];
            const val = String(nextRow[idx] || '').trim();
            if (val) {
                // 현재 부서 계층에 명시적인 값이 들어왔으므로, 그보다 하위 부서의 carry는 모두 초기화 (수평 병합 대비)
                for (let j = i + 1; j < deptCols.length; j++) {
                    delete carry[deptCols[j]];
                }
            }
        }

        uniqueFillable.forEach(index => {
            const value = String(nextRow[index] || '').trim();
            if (value) {
                carry[index] = nextRow[index];
            } else if (carry[index]) {
                nextRow[index] = carry[index];
            }
        });

        // 3. 계층 구조 결합 (예: 기술영업실 > 기술영업그룹)
        if (deptCols.length > 1 && deptIdx >= 0) {
            const depts = deptCols.map(idx => String(nextRow[idx] || '').trim()).filter(Boolean);
            
            // 인접한 동일 부서명 중복 제거
            const uniqueDepts = [];
            depts.forEach(d => {
                if (uniqueDepts.length === 0 || uniqueDepts[uniqueDepts.length - 1] !== d) {
                    uniqueDepts.push(d);
                }
            });
            
            if (options.showLowestDepartmentOnly) {
                nextRow[deptIdx] = uniqueDepts.length > 0 ? uniqueDepts[uniqueDepts.length - 1] : '';
            } else {
                nextRow[deptIdx] = uniqueDepts.join(' > ');
            }
            
            // 하위 부서 컬럼의 값은 지워서 중복 인식 방지
            for (let i = 1; i < deptCols.length; i++) {
                nextRow[deptCols[i]] = '';
            }
        }

        return nextRow;
    });
}

function buildPerson_(record, mapping, displayName) {
    const name = String(getMappedValue_(record, mapping.name) || '').trim();
    const phone = normalizePhone_(getMappedValue_(record, mapping.phone));
    const email = String(getMappedValue_(record, mapping.email) || '').trim();
    const company = String(getMappedValue_(record, mapping.company) || '').trim();
    const department = String(getMappedValue_(record, mapping.department) || '').trim();
    const title = String(getMappedValue_(record, mapping.title) || '').trim();
    const note = String(getMappedValue_(record, mapping.note) || '').trim();

    const person = {
        names: [{ givenName: displayName }],
        phoneNumbers: [{ value: phone }],
    };

    if (email) person.emailAddresses = [{ value: email }];
    if (company || department || title) {
        person.organizations = [{
            name: company || undefined,
            department: department || undefined,
            title: title || undefined,
        }];
    }
    if (note) person.biographies = [{ value: note, contentType: 'TEXT_PLAIN' }];

    return person;
}

function findOrCreateContactGroup_(labelName) {
    const groups = [];
    let pageToken = null;

    do {
        const response = People.ContactGroups.list({
            pageSize: 1000,
            pageToken,
            groupFields: 'metadata,name,groupType',
        });
        (response.contactGroups || []).forEach(group => groups.push(group));
        pageToken = response.nextPageToken;
    } while (pageToken);

    const existing = groups.find(group =>
        group.groupType === 'USER_CONTACT_GROUP' &&
        String(group.name || '').trim() === labelName
    );
    if (existing) return existing;

    return People.ContactGroups.create({ contactGroup: { name: labelName } });
}

function clearContactGroupContacts_(contactGroupResourceName) {
    const group = People.ContactGroups.get(contactGroupResourceName, { maxMembers: 10000 });
    const members = group.memberResourceNames || [];
    let deleted = 0;

    members.forEach(resourceName => {
        try {
            People.People.deleteContact(resourceName);
            deleted += 1;
        } catch (error) {
            // 이미 삭제됐거나 접근할 수 없는 연락처는 전체 동기화를 중단하지 않습니다.
        }
    });

    return deleted;
}

function addContactsToGroup_(contactGroupResourceName, resourceNames) {
    for (let i = 0; i < resourceNames.length; i += CONTACT_GROUP_MEMBER_CHUNK_SIZE) {
        const chunk = resourceNames.slice(i, i + CONTACT_GROUP_MEMBER_CHUNK_SIZE);
        People.ContactGroups.Members.modify({ resourceNamesToAdd: chunk }, contactGroupResourceName);
    }
}

function renderTemplate_(template, record) {
    return String(template || '{이름}').replace(/\{([^}]+)\}/g, (_, token) => {
        const key = String(token).trim();
        return record[key] == null ? '' : String(record[key]).trim();
    }).replace(/\s+/g, ' ').trim();
}

function rowToRecord_(row, headers) {
    const record = {};
    headers.forEach((header, index) => {
        record[header] = row[index] == null ? '' : row[index];
    });
    return record;
}

function getMappedValue_(record, headerName) {
    if (!headerName) return '';
    return record[headerName] == null ? '' : record[headerName];
}

function validateRequiredMapping_(mapping) {
    if (!mapping.name || !mapping.phone) {
        throw new Error('필수 매핑인 이름과 전화번호를 선택해야 합니다.');
    }
}

function buildDefaultTemplate_(headers) {
    const mapping = detectMapping_(headers);
    const parts = [];
    if (mapping.company) parts.push(`{${mapping.company}}`);
    if (mapping.name) parts.push(`{${mapping.name}}`);
    if (mapping.title) parts.push(`{${mapping.title}}`);
    return parts.length ? parts.join(' ') : '{이름}';
}

function makeUniqueHeaders_(headers) {
    const counts = {};
    return headers.map((header, index) => {
        const base = String(header || `열${index + 1}`).trim() || `열${index + 1}`;
        counts[base] = (counts[base] || 0) + 1;
        return counts[base] === 1 ? base : `${base}_${counts[base]}`;
    });
}

function normalizeHeader_(value) {
    return String(value || '').toLowerCase().replace(/[\s_\-./()[\]]/g, '');
}

function normalizePhone_(value) {
    return String(value || '').trim();
}

function getUserSettings_() {
    const raw = PropertiesService.getUserProperties().getProperty(USER_SETTINGS_KEY);
    if (!raw) return {};
    try {
        return JSON.parse(raw);
    } catch (error) {
        return {};
    }
}
