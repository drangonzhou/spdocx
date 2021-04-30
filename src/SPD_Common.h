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
#include <atomic>

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

// <1> Error Code

enum SPD_Error_t 
{
	SPD_ERR_OK = 0,
	SPD_ERR_ERROR = -9000,
	SPD_ERR_OPEN_ZIP,
	SPD_ERR_OPEN_XML,

	SPD_ERR_MIN
};


// <2> Log

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


// <3> Reference Object

class SPD_API RefObj
{
protected:
	RefObj() : m_val( 0 ) { }
	virtual ~RefObj() { }

protected:
	template< class T > friend class RefPtr;
	void IncRef() { m_val.fetch_add( 1 ); return; }
	void DecRef()
	{
		int ret = m_val.fetch_sub( 1 );
		if( ret == 1 )
			delete this;
		return;
	}

private:
	std::atomic_int m_val;
};

template<class T> class RefPtr
{
public:
	RefPtr( T * obj ) { m_obj = obj; if( m_obj != nullptr ) m_obj->IncRef(); }
	RefPtr( const RefPtr & obj ) { m_obj = obj.m_obj; if( m_obj != nullptr ) m_obj->IncRef(); }
	~RefPtr() { if( m_obj != nullptr ) m_obj->DecRef(), m_obj = nullptr; }

	RefPtr & operator = ( const RefPtr & obj )
	{
		if( &obj == this )
			return *this;
		if( m_obj != nullptr )
			m_obj->DecRef();
		m_obj = obj.m_obj;
		if( m_obj != nullptr ) 
			m_obj->IncRef();
		return *this;
	}

	inline T * operator -> () { return m_obj; }

	inline operator bool() const { return m_obj != nullptr; }
	inline bool IsValid() const { return m_obj != nullptr; }

	template< class T2 >
	RefPtr( const RefPtr<T2> & obj ) { m_obj = static_cast<T *>( obj.m_obj ); if( m_obj != nullptr ) m_obj->IncRef();	}
	template< class T2 >
	RefPtr & operator = ( const RefPtr<T2> & obj )
	{
		if( &obj == this )
			return *this;
		if( m_obj != nullptr )
			m_obj->DecRef();
		m_obj = static_cast<T *>( obj.m_obj );
		if( m_obj != nullptr ) 
			m_obj->IncRef();
		return *this;
	}

private:
	template< class T2 > friend class RefPtr;
	T * m_obj;
};

// DLL export template 
template SPD_API class RefPtr<RefObj>;

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_COMMON_H
