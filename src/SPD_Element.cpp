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

// DLL export template 
template SPD_API class RefPtr<Element>;
template SPD_API class RefPtr<Paragraph>;
template SPD_API class RefPtr<Hyperlink>;
template SPD_API class RefPtr<Run>;
template SPD_API class RefPtr<Table>;
template SPD_API class RefPtr<TRow>;
template SPD_API class RefPtr<TCell>;

static bool is_skipped_node( pugi::xml_node nd )
{
	// don't skip unknown node, only skip known unused node and nodes that already used by parent
	if( nd.type() != pugi::node_element
		|| strcmp( nd.name(), "w:t" ) == 0
		|| strcmp( nd.name(), "w:pPr" ) == 0
		|| strcmp( nd.name(), "w:rPr" ) == 0
		|| strcmp( nd.name(), "w:tblPr" ) == 0
		|| strcmp( nd.name(), "w:tblGrid" ) == 0
		|| strcmp( nd.name(), "w:trPr" ) == 0
		|| strcmp( nd.name(), "w:tcPr" ) == 0
		|| strcmp( nd.name(), "w:sectPr" ) == 0
		) 
	{
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
		if( !nd.child( "w:t" ).empty() ) {
			ele = new Run( doc, nd );
		}
		// TODO (later) : other w:r, ex w:commentReference
		else {
			ele = new Element( ElementType::ELEMENT_TYPE_UNKNOWN, doc, nd );
		}
	}
	else if( strcmp( nd.name(), "w:hyperlink" ) == 0 ) {
		ele = new Hyperlink( doc, nd );
	}
	else if( strcmp( nd.name(), "w:tbl" ) == 0 ) {
		ele = new Table( doc, nd );
	}
	else if( strcmp( nd.name(), "w:tr" ) == 0 ) {
		ele = new TRow( doc, nd );
	}
	else if( strcmp( nd.name(), "w:tc" ) == 0 ) {
		ele = new TCell( doc, nd );
	}
	// TODO (later) : other known element
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

RefPtr<Element> Element::GetParent() const
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.parent();
	// if parent is "/w:document/w:body" , then treat it as invalid
	if( pnd.parent().parent() == m_nd.root() )
		return CreateElement( nullptr );
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetPrev() const
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.previous_sibling();
	while( ! pnd.empty() && is_skipped_node( pnd ) )
		pnd = pnd.previous_sibling();
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetNext() const
{
	if( m_type == ElementType::ELEMENT_TYPE_INVALID || m_doc == nullptr || m_nd.empty() )
		return CreateElement( nullptr );
	pugi::xml_node pnd = m_nd.next_sibling();
	while( ! pnd.empty() && is_skipped_node( pnd ) )
		pnd = pnd.next_sibling();
	return CreateElement( m_doc, pnd );
}

RefPtr<Element> Element::GetFirstChild() const
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
	, m_style_name( nullptr )
	, m_text( nullptr )
{

}

Paragraph::~Paragraph()
{
	ResetCache();
}

void Paragraph::ResetCache()
{
	m_style_name = nullptr;
	if( m_text != nullptr )
		delete m_text, m_text = nullptr;
	return;
}

const char * Paragraph::GetStyleName()
{
	if( m_style_name != nullptr )
		return m_style_name;
	m_style_name = "";
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	m_style_name = m_doc->get_style_name( attr.name() );
	return m_style_name;
}

const char * Paragraph::GetText()
{
	if( m_text != nullptr )
		return m_text->c_str();
	m_text = new std::string( "" );
	for( pugi::xml_node cnd = m_nd.first_child(); !cnd.empty(); cnd = cnd.next_sibling() ) {
		if( strcmp( cnd.name(), "w:r" ) == 0 ) {
			m_text->append( cnd.child( "w:t" ).text().get() );
		}
		else if( strcmp( cnd.name(), "w:hyperlink" ) == 0 ) {
			for( pugi::xml_node cnd2 = cnd.first_child(); !cnd2.empty(); cnd2 = cnd2.next_sibling() ) {
				if( strcmp( cnd2.name(), "w:r" ) == 0 ) {
					m_text->append( cnd2.child( "w:t" ).text().get() );
				}
			}
		}
	}
	return m_text->c_str();
}

Hyperlink::Hyperlink( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_HYPERLINK, doc, nd )
	, m_link_type( nullptr )
	, m_targetMode( nullptr )
	, m_target( nullptr )
	, m_text( nullptr )
{

}

Hyperlink::~Hyperlink()
{
	ResetCache();
}

void Hyperlink::ResetCache()
{
	m_link_type = nullptr;
	m_targetMode = nullptr;
	m_target = nullptr;
	if( m_text != nullptr )
		delete m_text, m_text = nullptr;
	return;
}

const char * Hyperlink::GetAnchor()
{
	return m_nd.attribute( "w:anchor" ).value();
}

const char * Hyperlink::GetLinkType()
{
	if( m_link_type != nullptr )
		return m_link_type;
	m_link_type = "";
	const Relationship * rela = m_doc->get_relationship( m_nd.attribute( "r:id" ).value() );
	if( rela != nullptr ) {
		m_link_type = rela->m_type.c_str();
	}
	return m_link_type;
}

const char * Hyperlink::GetTargetMode()
{
	if( m_targetMode != nullptr )
		return m_targetMode;
	m_targetMode = "";
	const Relationship * rela = m_doc->get_relationship( m_nd.attribute( "r:id" ).value() );
	if( rela != nullptr ) {
		m_targetMode = rela->m_targetMode.c_str();
	}
	return m_targetMode;
}

const char * Hyperlink::GetTarget()
{
	if( m_target != nullptr )
		return m_target;
	m_target = "";
	const Relationship * rela = m_doc->get_relationship( m_nd.attribute( "r:id" ).value() );
	if( rela != nullptr ) {
		m_target = rela->m_target.c_str();
	}
	return m_target;
}

const char * Hyperlink::GetText()
{
	if( m_text != nullptr )
		return m_text->c_str();
	m_text = new std::string( "" );
	for( pugi::xml_node cnd = m_nd.first_child(); !cnd.empty(); cnd = cnd.next_sibling() ) {
		if( strcmp( cnd.name(), "w:r" ) == 0 ) {
			m_text->append( cnd.child( "w:t" ).text().get() );
		}
	}
	return m_text->c_str();
}

Run::Run( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_RUN, doc, nd )
	, m_text( nullptr )
{
	
}

Run::~Run()
{
	ResetCache();
}

void Run::ResetCache()
{
	if( m_text != nullptr )
		delete m_text, m_text = nullptr;
	return;
}

const char * Run::GetColor()
{
	return m_nd.child( "w:rPr" ).child( "w:color" ).attribute( "w:val" ).value();
}

const char * Run::GetHighline()
{
	return m_nd.child( "w:rPr" ).child( "w:highlight" ).attribute( "w:val" ).value();
}

bool Run::GetBold()
{
	return !m_nd.child( "w:rPr" ).child( "w:b" ).empty();
}

bool Run::GetItalic()
{
	return !m_nd.child( "w:rPr" ).child( "w:i" ).empty();
}

const char * Run::GetUnderline()
{
	return m_nd.child( "w:rPr" ).child( "w:u" ).attribute( "w:val" ).value();
}

bool Run::GetStrike()
{
	return !m_nd.child( "w:rPr" ).child( "w:strike" ).empty();
}

bool Run::GetDoubleStrike()
{
	return !m_nd.child( "w:rPr" ).child( "w:dstrike" ).empty();
}

const char * Run::GetText()
{
	if( m_text != nullptr )
		return m_text->c_str();
	m_text = new std::string( "" );
	m_text->append( m_nd.child( "w:t" ).text().get() );
	return m_text->c_str();
}

Table::Table( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_TABLE, doc, nd )
	, m_rowNum( -1 )
{

}

Table::~Table()
{
	ResetCache();
}

void Table::ResetCache()
{
	m_rowNum = -1;
	m_colWidth.clear();
	return;
}

int Table::GetRowNum()
{
	if( m_rowNum > 0 )
		return m_rowNum;
	m_rowNum = 0;
	for( pugi::xml_node pnd = m_nd.child( "w:tr" ); !pnd.empty(); pnd = pnd.next_sibling( "w:tr" ) ) {
		++m_rowNum;
	}
	return m_rowNum;
}

int Table::GetColNum()
{
	if( (int)m_colWidth.size() > 0 )
		return (int)m_colWidth.size();
	m_colWidth.reserve( 8 );
	pugi::xml_node pnd = m_nd.child( "w:tblGrid" );
	if( m_nd.empty() )
		return -1;
	for( pnd = pnd.child( "w:gridCol" ); !pnd.empty(); pnd = pnd.next_sibling( "w:gridCol" ) ) {
		m_colWidth.push_back( pnd.attribute( "w:w" ).as_int() );
	}
	return (int)m_colWidth.size();
}

int Table::GetColWidth( int idx )
{
	int colnum = GetColNum();
	if( idx < 0 || idx >= colnum )
		return -1;
	return m_colWidth[idx];
}

TRow::TRow( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_TABLE_TR, doc, nd )
{
}

TRow::~TRow()
{

}

TCell::TCell( Document * doc, pugi::xml_node nd )
	: Element( ElementType::ELEMENT_TYPE_TABLE_TC, doc, nd )
	, m_span_num( -1 )
	, m_vmerge_type( VMergeType::VMERGE_INVALID )
	, m_text( nullptr )
{

}

TCell::~TCell()
{
	ResetCache();
}

void TCell::ResetCache()
{
	m_span_num = -1;
	m_vmerge_type = VMergeType::VMERGE_INVALID;
	if( m_text != nullptr )
		delete m_text, m_text = nullptr;
	return;
}

int TCell::GetSpanNum()
{
	if( m_span_num != -1 )
		return m_span_num;
	m_span_num = 0;
	pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:gridSpan" );
	if( !pnd.empty() )
		m_span_num = pnd.attribute( "w:val" ).as_int();
	return m_span_num;
}

VMergeType TCell::GetVMergeType()
{
	if( m_vmerge_type != VMergeType::VMERGE_INVALID )
		return m_vmerge_type;
	pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:vMerge" );
	if( !pnd.empty() ) {
		pugi::xml_attribute attr = pnd.attribute( "w:val" );
		if( !attr.empty() && strcmp( attr.value(), "restart" ) == 0 )
			m_vmerge_type = VMergeType::VMERGE_START;
		else
			m_vmerge_type = VMergeType::VMERGE_CONT;
	}
	else {
		m_vmerge_type = VMergeType::VMERGE_NONE;
	}
	return m_vmerge_type;
}

const char * TCell::GetText()
{
	if( m_text != nullptr )
		return m_text->c_str();
	m_text = new std::string( "" );
	bool is_first = true;
	// only get paragraph text, ignore table in table
	for( RefPtr<Element> ele = GetFirstChild(); ele->IsValid(); ele = ele->GetNext() ) {
		if( ele->GetType() != ElementType::ELEMENT_TYPE_PARAGRAPH )
			continue;
		RefPtr<Paragraph> par = ele;
		if( is_first ) {
			is_first = false;
		}
		else {
			m_text->append( "\n" );
		}
		m_text->append( par->GetText() );
	}
	return m_text->c_str();
}

////////////////////////////////
END_NS_SPD
