#pragma once

#include "Settings.h"

struct AppSettings : Settings {
	BEGIN_SETTINGS(AppSettings)
		SETTING(MainWindowPlacement, WINDOWPLACEMENT{}, SettingType::Binary);
		SETTING(AlwaysOnTop, 0, SettingType::Bool);
		SETTING(SingleInstance, 0, SettingType::Bool);
		SETTING(ViewSystemClasses, 0, SettingType::Bool);
		SETTING(ViewSystemProperties, 0, SettingType::Bool);
		SETTING(ShowNamespacesInList, 0, SettingType::Bool);
	END_SETTINGS

	DEF_SETTING(AlwaysOnTop, int)
	DEF_SETTING(MainWindowPlacement, WINDOWPLACEMENT)
	DEF_SETTING(SingleInstance, int)
	DEF_SETTING(ViewSystemClasses, int)
	DEF_SETTING(ViewSystemProperties, int)
	DEF_SETTING(ShowNamespacesInList, int)
};
