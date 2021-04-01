// SPD_Def.h : spdocx common define
// Copyright (C) 2021 ~ 2021 drangon <drangon_zhou (at) hotmail.com>
//
// This program is free software: you can redistribute it and/or modify it
// under the terms of the GNU Lesser General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public License
// along with this program; if not, see <http://www.gnu.org/licenses/>.

#ifndef INCLUDED_SPD_COMMON_H
#define INCLUDED_SPD_COMMON_H

#include <stdarg.h>

#ifdef _WIN32
#ifdef LIBSPDOCX_EXPORTS
#define SPD_API __declspec(dllexport)
#else
#define SPD_API __declspec(dllimport)
#endif // BUILD_SPD_DLL
#else
#define SPD_API
#endif // _WIN32 else

#define BEGIN_NS_SPD namespace spd {
#define END_NS_SPD }

BEGIN_NS_SPD
////////////////////////////////

enum SPD_Error_t 
{
	SPD_ERR_OK = 0,
	SPD_ERR_ERROR = -9000,
	SPD_ERR_OPEN_ZIP,
	SPD_ERR_OPEN_XML,

	SPD_ERR_MIN
};

enum {
	SPD_LOG_LEVEL_DEBUG = 2,
	SPD_LOG_LEVEL_INFO = 4,
	SPD_LOG_LEVEL_ERR = 6,

	SPD_LOG_LEVEL_SILENCE = 99,
};

#ifdef _WIN32
#define SPD_LOG_LINE_END "\r\n"
#else
#define SPD_LOG_LINE_END "\n"
#endif

#define SPD_PR_DEBUG( fmt, ... ) SPD_PR_LOG( SPD_LOG_LEVEL_DEBUG, fmt, ##__VA_ARGS__ )
#define SPD_PR_INFO( fmt, ... ) SPD_PR_LOG( SPD_LOG_LEVEL_INFO, fmt, ##__VA_ARGS__ )
#define SPD_PR_ERR( fmt, ... ) SPD_PR_LOG( SPD_LOG_LEVEL_ERR, fmt, ##__VA_ARGS__ )
#define SPD_PR_LOG( level, fmt, ... ) spd::SPD_PrintLog( level, __FILE__, __LINE__, fmt SPD_LOG_LINE_END, ##__VA_ARGS__ )

SPD_API void SPD_PrintLog( int level, const char * file, int line, const char * fmt, ... );

typedef void ( *SPD_LogFunc_t )( const char * log_msg );

SPD_API int SPD_SetLogFuncLevel( SPD_LogFunc_t func, int level );

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_COMMON_H
