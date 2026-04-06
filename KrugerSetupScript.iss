; ==============================================================================
; Kruger Lab System - Inno Setup Script v2.0
; نظام إدارة المعمل - سكريبت التثبيت
; ==============================================================================
;
; الغرض: إنشاء مثبت تنفيذي لبرنامج Kruger Lab System
; الإصدار: 2.0.0
; التاريخ: 2026-04-06
; المطور: Kruger Lab Team
;
; هذا الملف يستخدم Inno Setup Compiler لإنشاء مثبت Windows
; لتطبيق إدارة المعمل المتكامل.
;
; المتطلبات:
; - Inno Setup Compiler (https://jrsoftware.org/isinfo.php)
; - ملفات البرنامج المبنية بـ PyInstaller
; - ملفات الإعدادات والقوالب
;
; طريقة الاستخدام:
; 1. تأكد من وجود ملف dist\Kruger_Lab_System.exe
; 2. افتح هذا الملف بـ Inno Setup Compiler
; 3. اضغط Build -> Compile
; 4. ستجد الملف المثبت في مجلد Output\
;
; ==============================================================================

[Setup]
; ==============================================================================
; قسم إعدادات التثبيت الأساسية
; ==============================================================================

; معلومات البرنامج - تم التحديث للإصدار 2.0
AppName=Kruger Lab System                    ; اسم البرنامج كما يظهر للمستخدم
AppVersion=2.0                               ; رقم الإصدار (يظهر في إضافة/إزالة البرامج)
AppVerName=Kruger Lab System 2.0             ; الاسم الكامل مع رقم الإصدار
AppPublisher=Kruger Lab                      ; اسم الناشر
AppPublisherURL=https://www.krugerlab.com    ; رابط موقع الناشر
AppSupportURL=https://www.krugerlab.com/support ; رابط الدعم الفني
AppUpdatesURL=https://www.krugerlab.com/updates ; رابط التحديثات

; إعدادات المجلدات والملفات
DefaultDirName={pf}\Kruger Lab System        ; مجلد التثبيت الافتراضي (Program Files)
DefaultGroupName=Kruger Lab System           ; اسم المجموعة في قائمة Start

; إعدادات الملف المثبت
OutputDir=Output                             ; مجلد حفظ الملف المثبت النهائي
OutputBaseFilename=Kruger_Lab_Setup_v2.0    ; اسم الملف المثبت
SetupIconFile=Setting\Logo.ico               ; أيقونة المثبت

; إعدادات الضغط
Compression=lzma                            ; نوع الضغط (LZMA هو الأفضل)
SolidCompression=yes                        ; ضغط متصل لتقليل الحجم

; صلاحيات النظام
PrivilegesRequired=admin                    ; طلب صلاحيات المسؤول (لإنشاء مجلدات في D:\)

; معلومات إضافية لإضافة/إزالة البرامج
VersionInfoVersion=2.0.0.0                  ; رقم الإصدار التقني
VersionInfoCompany=Kruger Lab               ; اسم الشركة
VersionInfoDescription=Kruger Lab Management System ; وصف البرنامج
VersionInfoCopyright=Copyright (C) 2024 Kruger Lab ; حقوق النشر

; إعدادات سلوك المثبت
AllowNoIcons=yes                            ; السماح بعدم إنشاء أيقونات
AlwaysRestart=no                            ; عدم إجبار إعادة التشغيل
AlwaysUsePersonalGroup=no                   ; عدم استخدام مجموعة شخصية دائماً
CreateUninstallRegKey=yes                   ; إنشاء مفتاح إلغاء التثبيت في السجل
UninstallDisplayIcon={app}\Kruger_Lab_System.exe ; أيقونة إلغاء التثبيت
UninstallDisplayName=Kruger Lab System 2.0  ; اسم إلغاء التثبيت

; ==============================================================================
; قسم اللغات المدعومة
; ==============================================================================
; يحدد اللغات المتاحة في المثبت
; المستخدم يمكنه اختيار اللغة أثناء التثبيت

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"     ; اللغة الإنجليزية (افتراضية)
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl" ; اللغة العربية

; ==============================================================================
; قسم المهام الاختيارية
; ==============================================================================
; يحدد الخيارات الإضافية المتاحة للمستخدم أثناء التثبيت
; كل مهمة يمكن تفعيلها أو تعطيلها حسب رغبة المستخدم

[Tasks]
; إنشاء أيقونة على سطح المكتب
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

; إنشاء أيقونة في شريط التشغيل السريع (للإصدارات القديمة من Windows فقط)
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; OnlyBelowVersion: 0,6.1

; ربط ملفات Excel بالبرنامج (جعل البرنامج يفتح تلقائياً عند النقر على ملفات .xlsx)
Name: "associate"; Description: "ربط ملفات Excel بالبرنامج"; GroupDescription: "ربط الملفات:"; Flags: unchecked

; ==============================================================================
; قسم إعدادات السجل (Registry)
; ==============================================================================
; يحدد التعديلات على سجل Windows لربط الملفات وإعدادات النظام
; هذه التعديلات تكون اختيارية حسب المهام المختارة

[Registry]
; ربط ملفات Excel بالبرنامج (اختياري - يعمل فقط إذا تم اختيار مهمة "associate")

; تحديد نوع الملف .xlsx ليستخدم نوع "KrugerLabExcel"
Root: HKCR; Subkey: ".xlsx"; ValueType: string; ValueName: ""; ValueData: "KrugerLabExcel"; Flags: uninsdeletevalue; Tasks: associate

; تعريف نوع الملف "KrugerLabExcel"
Root: HKCR; Subkey: "KrugerLabExcel"; ValueType: string; ValueName: ""; ValueData: "ملف Kruger Lab Excel"; Flags: uninsdeletekey; Tasks: associate

; تحديد الأيقونة للملفات من نوع KrugerLabExcel
Root: HKCR; Subkey: "KrugerLabExcel\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\Kruger_Lab_System.exe,0"; Tasks: associate

; تحديد الأمر الذي يتم تنفيذه عند فتح الملف
Root: HKCR; Subkey: "KrugerLabExcel\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\Kruger_Lab_System.exe"" ""%1"""; Tasks: associate

[UninstallDelete]
; ==============================================================================
; قسم تنظيف الملفات عند إلغاء التثبيت
; ==============================================================================
; يحدد الملفات والمجلدات التي سيتم حذفها عند إلغاء تثبيت البرنامج
; هذا يضمن عدم ترك ملفات متروكة على النظام

; حذف مجلد البرنامج بالكامل
Type: filesandordirs; Name: "{app}"

; حذف الملفات المؤقتة في مجلد المستندات
Type: files; Name: "{userdocs}\Kruger Lab System\*"

; ==============================================================================
; نهاية ملف التثبيت
; ==============================================================================
; هذا الملف يحتوي على جميع إعدادات التثبيت لبرنامج Kruger Lab System الإصدار 2.0
; لإنشاء ملف التثبيت، قم بتشغيل Inno Setup Compiler مع هذا الملف
; تأكد من وجود ملف Kruger_Lab_System.exe في مجلد dist قبل التجميع
; ==============================================================================

; ==============================================================================
; قسم المجلدات المراد إنشاؤها
; ==============================================================================
; يحدد المجلدات التي سيتم إنشاؤها تلقائياً أثناء التثبيت
; هذه المجلدات ضرورية لعمل البرنامج بشكل صحيح

[Directories]
; إنشاء شجرة المجلدات في قرص D تلقائياً عند التثبيت

; المجلد الرئيسي للبرنامج
Name: "D:\Plant Reports"

; مجلد الإعدادات والقوالب
Name: "D:\Plant Reports\Setting"

; مجلد التقارير اليومية
Name: "D:\Plant Reports\Daily Reports"
Name: "D:\Plant Reports\Daily Reports\Daily Sheets"    ; ملفات Excel اليومية
Name: "D:\Plant Reports\Daily Reports\Daily PDFs"      ; ملفات PDF المحولة

; مجلدات أخرى للبرنامج
Name: "D:\Plant Reports\Quality"                       ; تقارير الجودة
Name: "D:\Plant Reports\Dry Sludge"                    ; اختبارات الحمأة الجافة
Name: "D:\Plant Reports\External Plants"               ; تقارير المحطات الخارجية
Name: "D:\Plant Reports\Monthly Report"                ; التقارير الشهرية

; مجلدات جديدة للإصدار 2.0
Name: "D:\Plant Reports\Backup"                        ; النسخ الاحتياطية
Name: "D:\Plant Reports\Temp"                          ; الملفات المؤقتة
Name: "D:\Plant Reports\Logs"                          ; ملفات السجل

[Files]
; ==============================================================================
; قسم الملفات المراد نسخها
; ==============================================================================
; يحدد الملفات التي سيتم نسخها من مجلد التثبيت إلى مجلد البرنامج
; يشمل الملفات التنفيذية والإعدادات والقوالب

; ملف البرنامج الرئيسي (يجب أن تكون قد قمت بعمل الـ EXE أولاً باستخدام PyInstaller)
Source: "dist\Kruger_Lab_System.exe"; DestDir: "{app}"; Flags: ignoreversion

; نقل ملفات الإعدادات والقوالب لتعمل مباشرة من قرص D
Source: "Setting\*"; DestDir: "D:\Plant Reports\Setting"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; ==============================================================================
; قسم الأيقونات والاختصارات
; ==============================================================================
; يحدد الأيقونات والاختصارات التي سيتم إنشاؤها في قائمة Start وسطح المكتب
; هذه الاختصارات تسهل الوصول للبرنامج

; إنشاء اختصار في قائمة Start
Name: "{group}\Kruger Lab System"; Filename: "{app}\Kruger_Lab_System.exe"

; إنشاء اختصار على سطح المكتب (اختياري - يعمل فقط إذا تم اختيار مهمة desktopicon)
Name: "{commondesktop}\Kruger Lab System"; Filename: "{app}\Kruger_Lab_System.exe"; Tasks: desktopicon

; إنشاء اختصار في شريط التشغيل السريع (للإصدارات القديمة من Windows فقط)
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Kruger Lab System"; Filename: "{app}\Kruger_Lab_System.exe"; Tasks: quicklaunchicon

[Run]
; ==============================================================================
; قسم الأوامر المراد تنفيذها بعد التثبيت
; ==============================================================================
; يحدد الأوامر التي سيتم تنفيذها فور انتهاء التثبيت
; عادةً يستخدم لتشغيل البرنامج أو فتح الموقع أو إجراء إعدادات إضافية

; خيار تشغيل البرنامج فور انتهاء التثبيت
Filename: "{app}\Kruger_Lab_System.exe"; Description: "{cm:LaunchProgram,Kruger Lab System}"; Flags: nowait postinstall skipifsilent