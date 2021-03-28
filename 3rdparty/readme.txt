
[zlib]

(1) download zlib-1.2.11.zip from http://zlib.net/
(2) unzip zlib-1.2.11.zip
(3) use vs2019 to open contrib\vstudio\vc14\zlibvc.sln
(4) build zlibvc project, win32 and x64 release

[pugixml]

(1) download pugixml-1.11.zip from http://pugixml.org/
(2) unzip pugixml-1.11.zip
(3) add the following lines to pugiconfig.hpp

#ifdef PUGIXML_EXPORTS
#define PUGIXML_API __declspec(dllexport)
#else
#define PUGIXML_API __declspec(dllimport)
#endif

(4) create dll project, add PUGIXML_EXPORTS to predefined macro
(5) build dll project, win32 and x64 release
