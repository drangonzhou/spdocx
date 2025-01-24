
install tools : cmake-3.20


[zlib]

(1) download zlib-1.3.1.zip from http://zlib.net/
(2) unzip zlib-1.3.1.zip
(3) use vs2019 to open contrib\vstudio\vc17\zlibvc.sln
(4) toolchain change to v142, charset change to multibyte
(5) build zlibvc project, win32 and x64 release


[libzip]

(1) download libzip-1.11.3.zip from https://libzip.org/
(2) unzip zlib-1.11.3.tar.gz
(3) create bd dir and run : cmake -G "Visual Studio 16 2019" \
  -DZLIB_INCLUDE_DIR=path/to/3rdparty -DZLIB_LIBRARY=path/to/3rdparty/x64/zlibwapi.lib .. 
(4) charset change to multibyte
(5) build zip project, x64 release
(6) create bd2 dir and run : cmake -G "Visual Studio 16 2019" -A Win32 \
  -DZLIB_INCLUDE_DIR=path/to/3rdparty -DZLIB_LIBRARY=path/to/3rdparty/win32/zlibwapi.lib .. 
(7) charset change to multibyte, add ZLIB_WINAPI to predefine macro
(8) build zip project, win32 release


[pugixml]

(1) download pugixml-1.15.zip from http://pugixml.org/
(2) unzip pugixml-1.15.zip
(3) add the following lines to pugiconfig.hpp

#define PUGIXML_NO_EXCEPTIONS

#ifdef PUGIXML_EXPORTS
#define PUGIXML_API __declspec(dllexport)
#else
#define PUGIXML_API __declspec(dllimport)
#endif

(4) create dll project, add PUGIXML_EXPORTS to predefined macro, charset change to multibyte,
(5) build dll project, win32 and x64 release


