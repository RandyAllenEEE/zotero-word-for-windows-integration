
// Scans the document using a regex and converts matches to Zotero fields
statusCode __stdcall convertRegexToFields(document_t* doc, const wchar_t regexPattern[], int groupIndex, listNode_t** returnNode) {
	HANDLE_EXCEPTIONS_BEGIN
	
	*returnNode = NULL;
	
	// 1. Get full document text
	CRange contentRange = doc->comDoc.get_Content();
	CString cText = contentRange.get_Text();
	std::wstring docText((LPCTSTR)cText);

	// 2. Setup Regex
	// Word uses \r for paragraphs. ensure pattern handles it if user expects \n
	// But usually user provides pattern.
	std::wregex pattern(regexPattern);
	
	// 3. Iterate matches
	// We need to be careful about modifying the document while iterating indices.
	// However, we are replacing text with fields. The field code/result might differ in length from original text.
	// So we should work backwards? Or just calculate offsets?
	// If we work forwards, and replace text with a field, the indices of subsequent text shift.
	// "std::wsregex_iterator" works on the original string "docText".
	// So the "match.position()" is relative to the ORIGINAL text.
	// If we modify the document, the "start" in Word document shifts.
	// 
	// STRATEGY: Collect all matches first (ranges and keys), THEN apply changes from END to START to preserve indices.
	
	struct MatchInfo {
		long start;
		long length;
		std::wstring key;
	};
	
	std::vector<MatchInfo> matches;
	
	std::wsregex_iterator it(docText.begin(), docText.end(), pattern);
	std::wsregex_iterator end;

	for (; it != end; ++it) {
		std::wsmatch match = *it;
		if (groupIndex < match.size()) {
			MatchInfo info;
			info.start = match.position();
			info.length = match.length();
			info.key = match[groupIndex].str();
			matches.push_back(info);
		}
	}
	
	// 4. Apply changes (Reverse order)
	listNode_t* fieldListStart = NULL;
	listNode_t* fieldListEnd = NULL;
	
	setScreenUpdatingStatus(doc, false); // Optimize performance
	
	// Reserve vector to avoid realloc
	// Iterate backwards
	for (auto ri = matches.rbegin(); ri != matches.rend(); ++ri) {
		long start = ri->start; // 0-indexed from string
		long length = ri->length;
		
		// Word Range is 0-indexed? 
		// "Range.Start" properties returns character position.
		// "Content.Start" is usually 0.
		// We need to verify if Word treats index same as wstring.
		// Usually yes for simple text. usage of InsertFieldRaw will handle replacement.
		
		CRange range = doc->comDoc.Range(COleVariant((long)start), COleVariant((long)(start + length)));
		
		// Create field
		field_t* newField = NULL;
		// InsertFieldRaw deletes the content of range and inserts field
		statusCode status = insertFieldRaw(doc, L"Field", range, &newField);
		if (status == STATUS_OK && newField) {
			// Store Key in Code temporarily for JS to read
			// We use a prefix to identify it as a temp key if needed, or just raw key.
			// JS expects "ADDIN ZOTERO_ITEM ..." normally.
			// We can just put the key here.
			setCode(newField, ri->key.c_str());
			
			// Add to return list
			addValueToList(newField, &fieldListStart, &fieldListEnd);
		}
	}
	
	if (fieldListStart) {
        addValueToList(fieldListStart, &(doc->allocatedFieldListsStart), &(doc->allocatedFieldListsEnd));
        *returnNode = fieldListStart;
    }

	return STATUS_OK;
	HANDLE_EXCEPTIONS_END
}
