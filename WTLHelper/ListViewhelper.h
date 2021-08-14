#pragma once

struct ListViewHelper {
	static bool SaveAll(PCWSTR path, CListViewCtrl& lv, bool includeHeaders = true);
	static CString GetRowAsString(CListViewCtrl& lv, int row, WCHAR separator = L'\t');
	static int FindItem(CListViewCtrl& lv, PCWSTR text, bool partial);
};

