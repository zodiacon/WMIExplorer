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
		SETTING(DerivedInstances, 1, SettingType::Bool);
		SETTING(DarkMode, -1, SettingType::Int32);		// 1 dark, 0 light, -1 (never chosen): as the system is
	END_SETTINGS

	static constexpr PCWSTR RegistryKey = L"Software\\ScorpioSoftware\\WmiExp";

	DEF_SETTING(AlwaysOnTop, int)
	DEF_SETTING(MainWindowPlacement, WINDOWPLACEMENT)
	DEF_SETTING(SingleInstance, int)
	DEF_SETTING(ViewSystemClasses, int)
	DEF_SETTING(ViewSystemProperties, int)
	DEF_SETTING(ShowNamespacesInList, int)
	DEF_SETTING(DerivedInstances, int)
	DEF_SETTING(DarkMode, int)
};
