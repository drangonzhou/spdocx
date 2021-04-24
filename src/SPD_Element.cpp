// SPD_Element.cpp : spdocx docx element
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

#include "SPD_Element.h"
#include "SPD_Document.h"

BEGIN_NS_SPD
////////////////////////////////

static bool is_skipped_node( pugi::xml_node nd )
{
	if( nd.type() != pugi::node_element
		|| strcmp( nd.name(), "w:t" ) == 0
		|| strcmp( nd.name(), "w:pPr" ) == 0
		|| strcmp( nd.name(), "w:rPr" ) == 0
		|| strcmp( nd.name(), "w:tblPr" ) == 0
		|| strcmp( nd.name(), "w:tblGrid" ) == 0
		|| strcmp( nd.name(), "w:trPr" ) == 0
		|| strcmp( nd.name(), "w:tcPr" ) == 0
		|| strcmp( nd.name(), "w:sectPr" ) == 0
		) {
		return true;
	}
	return false;
}

RefPtr<Element> Element::CreateElement( Document * doc, pugi::xml_node nd )
{
	if( doc == nullptr || nd.empty() )
		return RefPtr<Element>( new Element( ElementType::ELEMENT_TYPE_INVALID, nullptr, pugi::xml_node() ) );

	Element * ele = nullptr;
	if( strcmp( nd.name(), "w:p" ) == 0 ) {
		ele = new Paragraph( doc, nd );
	}
	else if( strcmp( nd.name(), "w:r" ) == 0 ) {
		ele = new Run( doc, nd );
	}
	// TODO : other known element
	else {
		ele = new Element( ElementType::ELEMENT_TYPE_UNKNOWN, doc, nd );
	}

	return RefPtr<Element>( ele );
}

Element::Element( ElementType type, Document * doc, pugi::xml_node nd )
	: m_type( type ), m_doc( doc ), m_nd( nd )
{
	
}

Element::~Element()
{
	m_type = ElementType::ELEMENT_TYPE_INVALID;
	m_doc = nullptr;
	m_nd = pugi::xml_node();
}

RefPtr<Element> Element::GetParent()
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.parent();
	// if parent is "/w:document/w:body" , then treat it as invalid
	if( pnd.parent().parent() == m_nd.root() )
		return CreateElement( nullptr );
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetPrev()
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.previous_sibling();
	while( ! pnd.empty() && is_skipped_node( pnd ) )
		pnd = pnd.previous_sibling();
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetNext()
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.next_sibling();
	while( ! pnd.empty() && is_skipped_node( pnd ) )
		pnd = pnd.next_sibling();
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetFirstChild()
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_type == ElementType::ELEMENT_TYPE_UNKNOWN || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.first_child();
	while( ! pnd.empty() && is_skipped_node( pnd ) )
		pnd = pnd.next_sibling();
	return CreateElement( m_doc, pnd );
}

Paragraph::Paragraph( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_PARAGRAPH, doc, nd )
{
	// TODO : parse
}

Paragraph::~Paragraph()
{
}

Run::Run( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_RUN, doc, nd )
{
	// TODO : parse
}

Run::~Run()
{

}

////////////////////////////////
END_NS_SPD
