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
#include <string>
#include <vector>

BEGIN_NS_SPD
////////////////////////////////

enum class ElementTypeE : uint8_t
{
	INVALID,  // not valid
	UNKNOWN,  // valid but unknown tag

	PARAGRAPH,  // w:p (top)
	HYPERLINK,  // w:hyperlink
	RUN,        // w:r -> w:t

	TABLE,      // w:tbl (top)
	TABLE_ROW,   // w:tr
	TABLE_CELL,   // w:tc

	// TODO : picture

	SECTION,    // w:sectPr (top)

	BOOKMARK_START,   // w:bookmarkStart
	BOOKMARK_END,     // w:bookmarkEnd
	COMMENT_START,    // w:commentRangeStart
	COMMENT_END,      // w:commentRangeEnd
	RUN_COMMENT_REF,  // w:r -> w:commentReference

	MAX
};

class Relationship;
class Document;
class ElementIterator;
class ElementRange;

class SPD_API Element
{
public:
	Element() : m_type( ElementTypeE::INVALID ) { }
	Element( pugi::xml_node nd );
	Element( const Element & ele ) : m_nd( ele.m_nd ), m_type( ele.m_type ) { }
	~Element() { m_type = ElementTypeE::INVALID; }

	Element & operator = ( const Element & ele ) { if( &ele != this ) m_nd = ele.m_nd, m_type = ele.m_type; return *this; }

	bool IsValid() const { return m_type != ElementTypeE::INVALID; }
	operator bool() const { return IsValid(); }
	bool operator ! () const { return !IsValid(); }

	bool operator == ( const Element & ele ) const { return m_nd == ele.m_nd; }
	bool operator != ( const Element & ele ) const { return m_nd != ele.m_nd; }
	bool operator < ( const Element & ele ) const { return m_nd < ele.m_nd; }
	bool operator > ( const Element & ele ) const { return m_nd > ele.m_nd; }
	bool operator <= ( const Element & ele ) const { return m_nd <= ele.m_nd; }
	bool operator >= ( const Element & ele ) const { return m_nd >= ele.m_nd; }

	ElementTypeE GetType() const { return m_type; }
	const char * GetTag() const { return m_nd.name(); }

	Element GetParent() const;
	Element GetPrev() const;
	Element GetNext() const;
	Element GetFirstChild() const;

	// child
	inline ElementIterator begin() const;
	inline ElementIterator end() const;
	inline ElementRange Children() const;

	// TODO : create next/prev element, delete element

protected:
	friend class ElementIterator;
	pugi::xml_node m_nd;

private:
	friend class SPDDebug;
	ElementTypeE m_type;
};

class SPD_API ElementIterator
{
public:
	ElementIterator() { }
	ElementIterator( const Element & ele ) : m_ele( ele ) { }

	bool operator == ( const ElementIterator & it ) const { return m_ele.m_nd == it.m_ele.m_nd; }
	bool operator != ( const ElementIterator & it ) const { return m_ele.m_nd != it.m_ele.m_nd; }

	const ElementIterator & operator ++ () { m_ele = m_ele.GetNext(); return *this; }
	ElementIterator operator ++ ( int ) { auto tmp = *this; m_ele = m_ele.GetNext(); return tmp; }

	const ElementIterator & operator -- () { m_ele = m_ele.GetPrev(); return *this; }
	ElementIterator operator -- ( int ) { auto tmp = *this; m_ele = m_ele.GetPrev(); return tmp; }

	const Element & operator * () const { return m_ele; }
	const Element * operator -> () const { return &m_ele; }
	Element & operator * () { return m_ele; }
	Element * operator -> () { return &m_ele; }

private:
	Element m_ele;
};

class SPD_API ElementRange
{
public:
	ElementRange( ElementIterator begin_it, ElementIterator end_it )
		: m_begin( begin_it ), m_end( end_it ) { }

	ElementIterator begin() const { return m_begin; }
	ElementIterator end() const { return m_end; }

private:
	ElementIterator m_begin;
	ElementIterator m_end;
};

inline ElementIterator Element::begin() const { return ElementIterator( GetFirstChild() ); }
inline ElementIterator Element::end() const { return ElementIterator(); }
inline ElementRange Element::Children() const { return ElementRange(begin(), end()); }

class SPD_API Paragraph : public Element
{
private:
	Paragraph( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~Paragraph() { }

public:
	const char * GetStyleId() const;
	const char * GetStyleName( const Document * doc ) const; // style name, ex : heading 1
	std::string GetText() const;

	// TODO : modify
};

class SPD_API Hyperlink : public Element
{
private:
	Hyperlink( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~Hyperlink() { }

public:
	const char * GetAnchor() const;     // local bookmark link
	const char * GetRelationshipId() const;
	const Relationship * GetRelationship( const Document * doc ) const;
	const char * GetLinkType( const Document * doc ) const;   // hyperlink type, ex : image -> http://schema.../image
	const char * GetTargetMode( const Document * doc ) const; // hyperlink mode, internal : "", external : "External"
	const char * GetTarget( const Document * doc ) const;     // hyperlink target, internal : "media/image1.png", external : "http://xxx.org"
	std::string GetText() const;

	// TODO : modify
};

class SPD_API Run : public Element
{
private:
	Run( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~Run() { }

public:
	const char * GetColor() const;     // font front color, ex : 00B0F0
	const char * GetHighline() const;  // font bg color, ex : yellow
	bool GetBold() const;              // font bold 
	bool GetItalic() const;            // font italic
	const char * GetUnderline() const; // font underline, ex : "", "singal", "double"
	bool GetStrike() const;            // font deleted with strike
	bool GetDoubleStrike() const;      // font deleted with double strike
	std::string GetText() const;

	// TODO : modify
	void SetText( const char * text );
};

class SPD_API Table : public Element
{
private:
	Table( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~Table() { }

public:
	int GetRowNum() const;
	int GetColNum() const;
	// get column width, total page width maybe 8000+, return -1 if bad idx
	std::vector<int> GetColWidth() const;
	int GetColWidth( int idx ) const;  

	// TODO : modify
};

class SPD_API TRow : public Element
{
private:
	TRow( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~TRow() { }

public:
	// TODO : modify

};

enum class VMergeTypeE : int8_t
{
	INVALID = -1,
	NONE = 0,      // no vmerge, normal cell
	START,
	CONT,
};

class SPD_API TCell : public Element
{
private:
	TCell( pugi::xml_node nd ) : Element( nd ) { }
	virtual ~TCell() { }

public:
	int GetSpanNum() const;  // 0 means no span
	VMergeTypeE GetVMergeType() const;
	std::string GetText() const;

	// TODO : modify
};

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_ELEMENT_H
