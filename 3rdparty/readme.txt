
install tools : cmake-3.20


[zlib 1.3.1]

(1) download zlib131.zip from http://zlib.net/
(2) unzip zlib131.zip
(3) use vs2019 to open contrib\vstudio\vc17\zlibvc.sln
(4) modify zlibvc project, toolchain change to v142, c++ standard to C++17,
  charset change to multibyte, debug version add 'd' to output name, 
  change zlibvc.def VERSION to 1.301
(5) build zlibvc project, win32 and x64 release
(6) copy zlib.h zconf.h zlibwapi.lib zlibwapi.dll to target dir


[libzip 1.11.3]

(1) download libzip-1.11.3.zip from https://libzip.org/
(2) unzip libzip-1.11.3.zip
(3) create bd dir and run : cmake -G "Visual Studio 16 2019" \
  -DZLIB_INCLUDE_DIR=path/to/3rdparty -DZLIB_LIBRARY=path/to/3rdparty/x64/zlibwapi.lib .. 
(4) open libzip.sln, modify zip project, c++ standard to C++17, charset change to multibyte,
(5) build zip project, x64 release
(6) copy zip.h zipconf.h zip.lib zip.dll to target dir
(7) create bd2 dir and run : cmake -G "Visual Studio 16 2019" -A Win32 \
  -DZLIB_INCLUDE_DIR=path/to/3rdparty -DZLIB_LIBRARY=path/to/3rdparty/win32/zlibwapi.lib .. 
(8) open libzip.sln, modify zip project, c++ standard to C++17, charset change to multibyte, 
  add ZLIB_WINAPI to predefine macro
(9) build zip project, win32 release
(10) copy zip.h zipconf.h zip.lib zip.dll to target dir


[pugixml 1.15]

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
  c++ standard to C++17,
(5) build dll project, win32 and x64 release, 
(6) copy pugiconfig.hpp pugixml.hpp pugixml.lib pugixml.dll to target dir


