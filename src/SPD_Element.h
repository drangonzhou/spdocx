// SPD_Element.h : spdocx docx element
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

#ifndef INCLUDED_SPD_ELEMENT_H
#define INCLUDED_SPD_ELEMENT_H

#include "SPD_Common.h"

#include <pugixml.hpp>
#include <atomic>

BEGIN_NS_SPD
////////////////////////////////

class RefObj
{
protected:
	RefObj() : m_val(0) { }
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
	RefPtr( T * obj ) { m_obj = obj; m_obj->IncRef(); }
	RefPtr( const RefPtr & obj ) { m_obj = obj.m_obj; m_obj->IncRef(); }
	~RefPtr() { m_obj->DecRef(); m_obj = nullptr; }

	RefPtr & operator = ( const RefPtr & obj ) 
	{
		if( &obj == this )
			return *this;
		if( m_obj != nullptr )
			m_obj->DecRef();
		m_obj = obj.m_obj;
		m_obj->IncRef();
		return * this;
	}

	T * operator -> () { return m_obj; }

private:
	T * m_obj;
};

class Document;

enum ElementType
{
	ELEMENT_TYPE_PARAGRAPH,
	ELEMENT_TYPE_TABLE,

	ELEMENT_TYPE_MAX
};

class Element : public RefObj
{
protected:
	Element( Document * doc, pugi::xml_node node );
	virtual ~Element();

public:
	ElementType Type();

protected:
	ElementType m_type;
};

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_ELEMENT_H
