
[zlib]

(1) download zlib-1.2.11.zip from http://zlib.net/
(2) unzip zlib-1.2.11.zip
(3) use vs2019 to open contrib\vstudio\vc14\zlibvc.sln
(4) build zlibvc project, win32 and x64 release


[pugixml]

(1) download pugixml-1.11.zip from http://pugixml.org/
(2) unzip pugixml-1.11.zip
(3) add the following lines to pugiconfig.hpp

#define PUGIXML_NO_EXCEPTIONS

#ifdef PUGIXML_EXPORTS
#define PUGIXML_API __declspec(dllexport)
#else
#define PUGIXML_API __declspec(dllimport)
#endif

(4) create dll project, add PUGIXML_EXPORTS to predefined macro
(5) build dll project, win32 and x64 release


[jsoncpp]

(1) download jsoncpp-1.9.4.zip from https://github.com/open-source-parsers/jsoncpp
(2) unzip jsoncpp-1.9.4.zip
(3) add the following lines to include/json/config.h

#define JSON_USE_EXCEPTION 0
#define JSON_DLL

(4) install cmake-3.20 , create bd dir and run : cmake -G "Visual Studio 16 2019" .. 
(5) build jsoncpp_lib prlject, x64 releasse
(6) create bd2 dir and run : cmake -G "Visual Studio 16 2019" -A Win32 .. 
(7) build jsoncpp_lib prlject, win32 releasse
