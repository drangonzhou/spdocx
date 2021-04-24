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

BEGIN_NS_SPD
////////////////////////////////

class Document;

enum class ElementType
{
	ELEMENT_TYPE_INVALID,  // not valid
	ELEMENT_TYPE_UNKNOWN,  // unknown tag

	ELEMENT_TYPE_PARAGRAPH,
	ELEMENT_TYPE_RUN,
	ELEMENT_TYPE_HYPERLINK,

	ELEMENT_TYPE_TABLE,
	ELEMENT_TYPE_TABLE_TR,
	ELEMENT_TYPE_TABLE_TC,

	ELEMENT_TYPE_SECTION,

	ELEMENT_TYPE_BOOKMARK_START,
	ELEMENT_TYPE_BOOKMARK_END,
	ELEMENT_TYPE_COMMENT_START,
	ELEMENT_TYPE_COMMENT_END,
	ELEMENT_TYPE_RUN_COMMENT_REF,

	ELEMENT_TYPE_MAX
};

class SPD_API Element : public RefObj
{
public:
	Element( ElementType type, Document * doc, pugi::xml_node nd );
	virtual ~Element();

	static RefPtr<Element> CreateElement( Document * doc, pugi::xml_node nd = pugi::xml_node() );

	ElementType GetType() const { return m_type; }
	const char * GetTag() const { return m_nd.name(); }

	RefPtr<Element> GetParent();
	RefPtr<Element> GetPrev();
	RefPtr<Element> GetNext();
	RefPtr<Element> GetFirstChild();

private:
	friend class SPDDebug;
	ElementType m_type;

protected:
	Document * m_doc;
	pugi::xml_node m_nd;
};

class SPD_API Paragraph : public Element
{
public:
	Paragraph( Document * doc, pugi::xml_node nd );
	virtual ~Paragraph();
	// TODO : func
};

class SPD_API Run : public Element
{
public:
	Run( Document * doc, pugi::xml_node nd );
	virtual ~Run();
	// TODO : func
};

/*
class SPD_API Table : public Element
{
public:
	Table( Document * doc, pugi::xml_node nd );
	virtual ~Table();
};
//*/

// DLL export template 
template SPD_API class RefPtr<Element>;
template SPD_API class RefPtr<Paragraph>;
template SPD_API class RefPtr<Run>;
//template SPD_API class RefPtr<Table>;

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_ELEMENT_H
