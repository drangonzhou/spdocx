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

#include "SPD_Common.h"

#include <windows.h>
#include <stdio.h>

BEGIN_NS_SPD
////////////////////////////////

// DLL export template 
template SPD_API class RefPtr<RefObj>;

static_assert( SPD_ERR_END < 0, "SPD_ERROR too much" );

static void spd_log_func_stderr( const char * log_msg )
{
	fputs( log_msg, stderr );
	return;
}

static SPD_LogFunc_t s_logfunc = spd_log_func_stderr;
static int s_level = SPD_LOG_LEVEL_INFO;

static const char * get_file_name( const char * file )
{
	const char * fname = file;
	for( const char * p = file; *p != '\0'; ++p ) {
		if( *p == '/' || *p == '\\' )
			fname = p + 1;
	}
	return fname;
}

SPD_API void SPD_PrintLog( int level, const char * file, int line, const char * fmt, ... )
{
	if( level < s_level )
		return;

	static const char * s_level_str[8] = { "0", "1", "DBG", "3", "INFO", "5", "ERR", "7" };
	char buf[2048];
	int len = 0;
	int ret = 0;

#ifdef _WIN32
	SYSTEMTIME mytm; 
	::GetLocalTime( &mytm );
	len = snprintf( buf, sizeof( buf ) - 1, "%02d%02d %02d:%02d:%02d.%06d %s:%d [%s] ",
		mytm.wMonth, mytm.wDay, mytm.wHour, mytm.wMinute, mytm.wSecond, mytm.wMilliseconds,
		get_file_name(file), line, s_level_str[level] );
#else
	struct timeval tv;
	gettimeofday( &tv, NULL );
	struct tm mytm;
	localtime_r( &tv, &mytm );
	len = snprintf( buf, sizeof( buf ) - 1, "%02d%02d %02d:%02d:%02d.%06d %s:%d [%s] ",
		mytm.tm_mon, mytm.tm_mday, mytm.tm_hour, mytm.tm_min, mytm.tm_sec, tv.tv_usec,
		get_file_name( file ), line, s_level_str[level] );
#endif

	va_list ap;
	va_start( ap, fmt );
	ret = vsnprintf( buf + len, sizeof( buf ) - len - 1, fmt, ap );
	va_end( ap );
	buf[sizeof( buf ) - 1] = '\0';
	if( ret >= 0 && ret <= (int)sizeof( buf ) - len - 1 )
		len += ret;
	else
		len = (int)strlen( buf );

	( *s_logfunc )( buf );
	return;
}

SPD_API int SPD_SetLogFuncLevel( SPD_LogFunc_t func, int level )
{
	s_logfunc = ( func == nullptr ) ? spd_log_func_stderr : func;
	if( level <= SPD_LOG_LEVEL_DEBUG )
		s_level = SPD_LOG_LEVEL_DEBUG;
	else if( level <= SPD_LOG_LEVEL_INFO )
		s_level = SPD_LOG_LEVEL_INFO;
	else if( level <= SPD_LOG_LEVEL_ERR )
		s_level = SPD_LOG_LEVEL_ERR;
	else
		s_level = SPD_LOG_LEVEL_SILENCE;
	return 0;
}

////////////////////////////////
END_NS_SPD
