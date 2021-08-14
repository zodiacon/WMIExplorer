#pragma once

#include "Settings.h"

struct AppSettings : Settings {
	BEGIN_SETTINGS(AppSettings)
		SETTING(MainWindowPlacement, WINDOWPLACEMENT{}, SettingType::Binary);
		SETTING(AlwaysOnTop, 0, SettingType::Bool);
		SETTING(SingleInstance, 0, SettingType::Bool);
	END_SETTINGS

	DEF_SETTING(AlwaysOnTop, int)
	DEF_SETTING(MainWindowPlacement, WINDOWPLACEMENT)
	DEF_SETTING(SingleInstance, int)
};
