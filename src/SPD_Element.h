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

class Document;

enum class ElementType
{
	ELEMENT_TYPE_INVALID,  // not valid
	ELEMENT_TYPE_UNKNOWN,  // unknown tag

	ELEMENT_TYPE_PARAGRAPH,  // w:p
	ELEMENT_TYPE_HYPERLINK,  // w:hyperlink
	ELEMENT_TYPE_RUN,        // w:r -> w:t

	ELEMENT_TYPE_TABLE,      // w:tbl
	ELEMENT_TYPE_TABLE_TR,   // w:tr
	ELEMENT_TYPE_TABLE_TC,   // w:tc

	// TODO : picture

	ELEMENT_TYPE_SECTION,    // w:sectPr

	ELEMENT_TYPE_BOOKMARK_START,   // w:bookmarkStart
	ELEMENT_TYPE_BOOKMARK_END,     // w:bookmarkEnd
	ELEMENT_TYPE_COMMENT_START,    // w:commentRangeStart
	ELEMENT_TYPE_COMMENT_END,      // w:commentRangeEnd
	ELEMENT_TYPE_RUN_COMMENT_REF,  // w:r -> w:commentReference

	ELEMENT_TYPE_MAX
};

class SPD_API Element : public RefObj
{
public:
	static RefPtr<Element> CreateElement( Document * doc, pugi::xml_node nd = pugi::xml_node() );

public:
	Element( ElementType type, Document * doc, pugi::xml_node nd );
	virtual ~Element();
	Element( const Element & ele ) = delete;
	Element & operator = ( const Element & ele ) = delete;

	virtual void ResetCache() { return; }

	bool IsValid() const { return m_type != ElementType::ELEMENT_TYPE_INVALID; }
	ElementType GetType() const { return m_type; }
	const char * GetTag() const { return m_nd.name(); }

	RefPtr<Element> GetParent() const;
	RefPtr<Element> GetPrev() const;
	RefPtr<Element> GetNext() const;
	RefPtr<Element> GetFirstChild() const;

	// TODO : create next/prev element, delete element

private:
	friend class SPDDebug;
	ElementType m_type;

protected:
	Document * m_doc;
	pugi::xml_node m_nd;
};

template class SPD_API RefPtr<Element>;

class SPD_API Paragraph : public Element
{
public:
	Paragraph( Document * doc, pugi::xml_node nd );
	virtual ~Paragraph();

	virtual void ResetCache();

	const char * GetStyleName(); // style name, ex : heading 1
	const char * GetText();

	// TODO : modify

protected:
	const char * m_style_name;
	std::string * m_text;
};

template class SPD_API RefPtr<Paragraph>;

class SPD_API Hyperlink : public Element
{
public:
	Hyperlink( Document * doc, pugi::xml_node nd );
	virtual ~Hyperlink();

	virtual void ResetCache();

	const char * GetAnchor();     // local bookmark link
	const char * GetLinkType();   // hyperlink type, ex : image -> http://schema.../image
	const char * GetTargetMode(); // hyperlink mode, internal : "", external : "External"
	const char * GetTarget();     // hyperlink target, internal : "media/image1.png", external : "http://xxx.org"
	const char * GetText();

	// TODO : modify

protected:
	const char * m_link_type;
	const char * m_target;
	const char * m_targetMode;
	std::string * m_text;
};

template class SPD_API RefPtr<Hyperlink>;

class SPD_API Run : public Element
{
public:
	Run( Document * doc, pugi::xml_node nd );
	virtual ~Run();

	virtual void ResetCache();

	const char * GetColor();     // font front color, ex : 00B0F0
	const char * GetHighline();  // font bg color, ex : yellow
	bool GetBold();              // font bold 
	bool GetItalic();            // font italic
	const char * GetUnderline(); // font underline, ex : "", "singal", "double"
	bool GetStrike();            // font deleted with strike
	bool GetDoubleStrike();      // font deleted with double strike
	const char * GetText();

	// TODO : modify
	void SetText( const char * text );

protected:
	std::string * m_text;
};

template class SPD_API RefPtr<Run>;

class SPD_API Table : public Element
{
public:
	Table( Document * doc, pugi::xml_node nd );
	virtual ~Table();

	virtual void ResetCache();

	int GetRowNum();
	int GetColNum();
	int GetColWidth( int idx );  // get each column width, total width maybe 8000+

	// TODO : modify

protected:
	int m_rowNum;
	std::vector<int> m_colWidth;
};

template class SPD_API RefPtr<Table>;

class SPD_API TRow : public Element
{
public:
	TRow( Document * doc, pugi::xml_node nd );
	virtual ~TRow();

	// TODO : modify

};

template class SPD_API RefPtr<TRow>;

enum class VMergeType
{
	VMERGE_INVALID = -1,
	VMERGE_NONE = 0,      // no vmerge, normal cell
	VMERGE_START,
	VMERGE_CONT,
};

class SPD_API TCell : public Element
{
public:
	TCell( Document * doc, pugi::xml_node nd );
	virtual ~TCell();

	virtual void ResetCache();

	int GetSpanNum();  // 0 means no span
	VMergeType GetVMergeType();
	const char * GetText();

	// TODO : modify

protected:
	int m_span_num;
	VMergeType m_vmerge_type;
	std::string * m_text;
};

template class SPD_API RefPtr<TCell>;

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_ELEMENT_H
