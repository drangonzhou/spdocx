// SPD_Element.h : spdocx docx element
// Copyright (C) 2021 ~ 2025 drangon <drangon_zhou (at) hotmail.com>
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
	RUN,        // w:r -> w:t or w:drawing or w:object

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

class Paragraph;
class Hyperlink;
class Run;
class Table; 
class TRow;
class TCell;

extern const std::string g_ElementEmptyStr;

class SPD_API Element
{
protected:
	friend class Document;
	Element( pugi::xml_node nd );

public:
	Element() : m_type( ElementTypeE::INVALID ) { }
	Element( const Element & ele ) : m_nd( ele.m_nd ), m_type( ele.m_type ) { }
	~Element() { m_type = ElementTypeE::INVALID; }

	static ElementTypeE GetNodeType( pugi::xml_node nd );

public:
	Element & operator = ( const Element & ele ) { if( &ele != this ) m_nd = ele.m_nd, m_type = ele.m_type; return *this; }

	bool IsValid() const { return (m_type != ElementTypeE::INVALID) && (! m_nd.empty()); }
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
	// auto skip some node ( w:t, w:pPr, w:rPr ... ), return INVALID if not found
	Element GetPrev() const;
	Element GetNext() const;
	Element GetFirstChild() const;

	Element GetPrev( ElementTypeE type ) const;
	Element GetNext( ElementTypeE type ) const;
	Element GetChild( ElementTypeE type ) const;

	// child
	inline ElementIterator begin() const;
	inline ElementIterator end() const;
	inline ElementRange Children() const;

	int DelChild( Element & child );
	int DelAllChild();

	// util
	static pugi::xml_node GetCreateChild( pugi::xml_node parent, const char * name );
	static pugi::xml_attribute GetCreateAttr( pugi::xml_node nd, const char * name );

protected:
	friend class ElementIterator;
	friend class Paragraph;
	friend class Hyperlink;
	friend class Run;
	friend class Table;
	friend class TRow;
	friend class TCell;
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

// NOTE : top level is Paragraph or Table

class SPD_API Paragraph : public Element
{
protected:
	friend class Document;
	Paragraph( pugi::xml_node nd ) : Element( nd ) { }

public:
	Paragraph( const Element & ele ) : Element( ele ) { }
	virtual ~Paragraph() { }

public:
	const char * GetStyleId() const;
	const char * GetStyleName( const Document & doc ) const; // style name, ex : heading 1
	const char * GetNumId() const;    // nullptr means no numid, use style default
	int GetNumLevel() const;          // -1 means no level, use style default
	std::string GetText() const;

	int SetStyleId( const char * id ); // id must valid in doc
	int SetStyleName( const char * name, const Document & doc );
	int SetNumId( const char * id ); // empty means remove numid, use style default
	int SetNumLevel( int level ); // use with SetNumId(), -1 means remove numlevel, use style default
	// Text can not set directly, need to set child run

	// NOTE : child is Run or Hyperlink ( or Bookmark or Comment )
	Run AddChildRun( bool add_back = true );
	Hyperlink AddChildHyperlink( bool add_back = true );

	// NOTE : sibling is Paragraph or Table
	Paragraph AddSiblingParagraph( bool add_next = true );
	Table AddSiblingTable( bool add_next = true );
};

class SPD_API Hyperlink : public Element
{
protected:
	friend class Document;
	Hyperlink( pugi::xml_node nd ) : Element( nd ) { }

public:
	Hyperlink( const Element & ele ) : Element( ele ) { }
	virtual ~Hyperlink() { }

public:
	const char * GetAnchor() const;          // local bookmark link
	const char * GetRelaId() const;  // remote hiperlink
	const Relationship * GetRela( const Document & doc ) const;
	const std::string & GetRelaType( const Document & doc ) const;   // hyperlink type, ex : hyperlink
	const std::string & GetRelaTargetMode( const Document & doc ) const; // hyperlink mode, internal : "", external : "External"
	const std::string & GetRelaTarget( const Document & doc ) const;     // hyperlink target, internal : "media/image1.png", external : "http://xxx.org"
	std::string GetText() const;

	int SetAnchor( const char * anchor );          // local bookmark link
	int SetRelaId( const char * relid );   // remote hiperlink
	// relationship detail and text need to set through Document

	// NOTE : child is Run
	Run AddChildRun( bool add_back = true );

	// NOTE : sibling is Hyperlink or Run ( or Bookmark or Comment )
	Hyperlink AddSiblingHyperlink( bool add_next = true );
	Run AddSiblingRun( bool add_next = true );
};

class SPD_API Run : public Element
{
protected:
	friend class Document;
	Run( pugi::xml_node nd ) : Element( nd ) { }

public:
	Run( const Element & ele ) : Element( ele ) { }
	virtual ~Run() { }

public:
	bool IsText() const;	 // text or not
	bool IsPic() const;	     // picture or not
	bool IsObject() const;	 // object or not

	const char * GetColor() const;     // font front color, ex : 00B0F0
	const char * GetHighline() const;  // font bg color, ex : yellow
	bool GetBold() const;              // font bold 
	bool GetItalic() const;            // font italic
	const char * GetUnderline() const; // font underline, ex : "", "single", "double", "dotted"
	bool GetStrike() const;            // font deleted with strike
	bool GetDoubleStrike() const;      // font deleted with double strike, can not both exist with strike
	std::string GetText() const;

	int SetColor( const char * color );
	int SetHighline( const char * color );
	int SetBold( bool bold = true );
	int SetItalic( bool italic = true );
	int SetUnderline( const char * underline = "singal" );
	int SetStrike( bool strike = true );
	int SetDoubleStrike( bool dstrike = true );
	void SetText( const char * text );

	const char * GetPicRelaId() const; // picture id
	int GetPicData( const Document & doc, std::vector<char> & data ) const;
	int SetPic( const char * id );
	int SetPic( const char * id, Document & doc, const std::vector<char> & data );
	int SetPic( const char * id, Document & doc, std::vector<char> && data );

	const char * GetObjectRelaId() const; // object id
	const char * GetObjectProgId() const; // object prog id
	const char * GetObjectImgRelaId() const; // object image id
	int GetObjectData( const Document & doc, std::vector<char> & data ) const;
	int GetObjectImgData( const Document & doc, std::vector<char> & data ) const;
	int SetObject( const char * objid, const char * progid, const char * imgid );
	int SetObject( const char * objid, const char * progid, const char * imgid, Document & doc, const std::vector<char> & objdata, const std::vector<char> & imgdata );
	int SetObject( const char * objid, const char * progid, const char * imgid, Document & doc, std::vector<char> && objdata, std::vector<char> && imgdata );

	// NOTE : no child ( w:t text is skip and handle by Run )

	// NOTE : sibling is Hyperlink or Run ( or Bookmark or Comment )
	Hyperlink AddSiblingHyperlink( bool add_next = true );
	Run AddSiblingRun( bool add_next = true );
};

// TODO : bookmark
// TODO : comment

class SPD_API Table : public Element
{
protected:
	friend class Document;
	Table( pugi::xml_node nd ) : Element( nd ) { }

public:
	Table( const Element & ele ) : Element( ele ) { }
	virtual ~Table() { }

public:
	int GetRowNum() const;
	int GetColNum() const;
	// get column width, total page width maybe 8100, return -1 if bad idx
	std::vector<int> GetColWidth() const;
	int GetColWidth( int idx ) const;  

	// NOTE : child is Row, if set Col, update all child
	int AddCol( int index, int num = 1 ); // insert new column at index, index begin from 0
	int DelCol( int index, int num = 1 ); // del column at index
	int SetColWidth( const std::vector<int> widths ); // min col width is 100, usually shound not < 300
	TRow AddChildTRow( bool add_back = true );
	int Reset( int row = 1, int col = 1 ); // row and col should >= 1, col should < 80;

	//
	int DelRow( Element & row ); // del row, row must exist, need special handle of Cell with VMerge

	// NOTE : sibling is Paragraph or Table
	Paragraph AddSiblingParagraph( bool add_next = true );
	Table AddSiblingTable( bool add_next = true );
};

class SPD_API TRow : public Element
{
private:
	friend class Document;
	TRow( pugi::xml_node nd ) : Element( nd ) { }

public: 
	TRow( const Element & ele ) : Element( ele ) { }
	virtual ~TRow() { }

public:
	TCell GetCell( int col, int & idx );

	// NOTE : child is Cell, but can not add/del directly, use Table AddCol/DelCol, or use TCell VMerge/Span
	
	// NOTE : sibling is TRow
	TRow AddSiblingTRow( bool add_next = true );
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
	friend class Document;
	TCell( pugi::xml_node nd ) : Element( nd ) { }

public:
	TCell( const Element & ele ) : Element( ele ) { }
	virtual ~TCell() { }

public:
	int GetCol() const;
	int GetSpanNum() const;       // 1 means no span, has span should >= 1
	VMergeTypeE GetVMergeType() const;
	int GetVMergeNum() const;     // 1 means no vmerge, vmerge start should >= 1, not valid for vmerge cont return -1
	std::string GetText() const;  // only paragraph, ignore table

	// NOTE : modify Span/VMerge only can modify first Cell in Span/VMerge
	int SetSpanNum( int num ); // 1 means no span, more span num will remove sibling TCell
	// 1 means no vmerge, more vmerge will set this cell START and more v-sibliing TCell CONT
	int SetVMergeType( VMergeTypeE type );  // note, need handle v-sibling TCell by hand if needed
	int SetVMergeNum( int num );            // 1 means no vmerge, more vmerge will set this cell START and more v-sibliing TCell CONT

	// NOTE : child is Paragraph or Table
	Paragraph AddChildParagraph( bool add_back = true );
	Table AddChildTable( bool add_back = true );

	// NOTE : can not add/del sibling TCell directly, use Table add/del col, or use TCell VMerge/Span
};

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_ELEMENT_H
