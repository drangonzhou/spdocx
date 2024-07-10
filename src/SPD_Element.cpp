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

Element::Element( pugi::xml_node nd ) : m_nd( nd )
{
	if( !nd ) {
		m_type = ElementTypeE::INVALID;
	}
	else if( strcmp( nd.name(), "w:p" ) == 0 ) {
		m_type = ElementTypeE::PARAGRAPH;
	}
	else if( strcmp( nd.name(), "w:r" ) == 0 ) {
		if( !nd.child( "w:t" ).empty() ) {
			m_type = ElementTypeE::RUN;
		}
		// TODO (later) : other w:r, ex w:commentReference
		else {
			m_type = ElementTypeE::UNKNOWN;
		}
	}
	else if( strcmp( nd.name(), "w:hyperlink" ) == 0 ) {
		m_type = ElementTypeE::HYPERLINK;
	}
	else if( strcmp( nd.name(), "w:tbl" ) == 0 ) {
		m_type = ElementTypeE::TABLE;
	}
	else if( strcmp( nd.name(), "w:tr" ) == 0 ) {
		m_type = ElementTypeE::TABLE_ROW;
	}
	else if( strcmp( nd.name(), "w:tc" ) == 0 ) {
		m_type = ElementTypeE::TABLE_CELL;
	}
	// TODO (later) : other known element
	else {
		m_type = ElementTypeE::UNKNOWN;
	}
	return;
}

Element Element::GetParent() const
{
	if( m_type == ElementTypeE::INVALID )
		return Element();
	pugi::xml_node pnd = m_nd.parent();
	// if parent is "/w:document/w:body" , then treat it as invalid
	if( pnd.parent().parent() == m_nd.root() )
		return Element();
	return Element( pnd );
}

Element Element::GetPrev() const
{
	if( m_type == ElementTypeE::INVALID )
		return Element();
	pugi::xml_node nd = m_nd.previous_sibling();
	while( ! nd.empty() && is_skipped_node( nd ) )
		nd = nd.previous_sibling();
	return Element( nd );
}

Element Element::GetNext() const
{
	if( m_type == ElementTypeE::INVALID )
		return Element();
	pugi::xml_node nd = m_nd.next_sibling();
	while( ! nd.empty() && is_skipped_node( nd ) )
		nd = nd.next_sibling();
	return Element( nd );
}

Element Element::GetFirstChild() const
{
	if( m_type == ElementTypeE::INVALID || m_type == ElementTypeE::UNKNOWN )
		return Element();
	pugi::xml_node nd = m_nd.first_child();
	while( ! nd.empty() && is_skipped_node( nd ) )
		nd = nd.next_sibling();
	return Element( nd );
}

const char * Paragraph::GetStyleId() const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	return attr.name();
}

const char * Paragraph::GetStyleName( const Document * doc ) const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	return doc->GetStyleName( attr.name() );
}

std::string Paragraph::GetText() const
{
	std::string text;
	for( pugi::xml_node cnd = m_nd.first_child(); !cnd.empty(); cnd = cnd.next_sibling() ) {
		if( strcmp( cnd.name(), "w:r" ) == 0 ) {
			text.append( cnd.child( "w:t" ).text().get() );
		}
		else if( strcmp( cnd.name(), "w:hyperlink" ) == 0 ) {
			for( pugi::xml_node cnd2 = cnd.first_child(); !cnd2.empty(); cnd2 = cnd2.next_sibling() ) {
				if( strcmp( cnd2.name(), "w:r" ) == 0 ) {
					text.append( cnd2.child( "w:t" ).text().get() );
				}
			}
		}
	}
	return text;
}

const char * Hyperlink::GetAnchor() const
{
	return m_nd.attribute( "w:anchor" ).value();
}

const char * Hyperlink::GetRelationshipId() const
{
	return m_nd.attribute( "r:id" ).value();
}

const Relationship * Hyperlink::GetRelationship( const Document * doc ) const
{
	return doc->GetRelationship( m_nd.attribute( "r:id" ).value() );
}

const char * Hyperlink::GetLinkType( const Document * doc ) const
{
	const Relationship * rela = doc->GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return nullptr;
	}
	return rela->m_type.c_str();
}

const char * Hyperlink::GetTargetMode( const Document * doc ) const
{
	const Relationship * rela = doc->GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return nullptr;
	}
	return rela->m_targetMode.c_str();
}

const char * Hyperlink::GetTarget( const Document * doc ) const
{
	const Relationship * rela = doc->GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return nullptr;
	}
	return rela->m_target.c_str();
}

std::string Hyperlink::GetText() const
{
	std::string text;
	for( pugi::xml_node cnd = m_nd.first_child(); !cnd.empty(); cnd = cnd.next_sibling() ) {
		if( strcmp( cnd.name(), "w:r" ) == 0 ) {
			text.append( cnd.child( "w:t" ).text().get() );
		}
	}
	return text;
}

const char * Run::GetColor() const
{
	return m_nd.child( "w:rPr" ).child( "w:color" ).attribute( "w:val" ).value();
}

const char * Run::GetHighline() const
{
	return m_nd.child( "w:rPr" ).child( "w:highlight" ).attribute( "w:val" ).value();
}

bool Run::GetBold() const
{
	return !m_nd.child( "w:rPr" ).child( "w:b" ).empty();
}

bool Run::GetItalic() const
{
	return !m_nd.child( "w:rPr" ).child( "w:i" ).empty();
}

const char * Run::GetUnderline() const
{
	return m_nd.child( "w:rPr" ).child( "w:u" ).attribute( "w:val" ).value();
}

bool Run::GetStrike() const
{
	return !m_nd.child( "w:rPr" ).child( "w:strike" ).empty();
}

bool Run::GetDoubleStrike() const
{
	return !m_nd.child( "w:rPr" ).child( "w:dstrike" ).empty();
}

std::string Run::GetText() const
{
	std::string text;
	text.append( m_nd.child( "w:t" ).text().get() );
	return text;
}

void Run::SetText( const char * text )
{
	m_nd.child( "w:t" ).text().set( text );
	return;
}

int Table::GetRowNum() const
{
	int rowNum = 0;
	for( pugi::xml_node pnd = m_nd.child( "w:tr" ); !pnd.empty(); pnd = pnd.next_sibling( "w:tr" ) ) {
		++rowNum;
	}
	return rowNum;
}

int Table::GetColNum() const
{
	pugi::xml_node pnd = m_nd.child( "w:tblGrid" );
	if( m_nd.empty() )
		return -1;
	int num = 0;
	for( pnd = pnd.child( "w:gridCol" ); !pnd.empty(); pnd = pnd.next_sibling( "w:gridCol" ) ) {
		++num;
	}
	return num;
}

std::vector<int> Table::GetColWidth() const
{
	std::vector<int> colWidth;
	colWidth.reserve( 8 );
	pugi::xml_node pnd = m_nd.child( "w:tblGrid" );
	if( m_nd.empty() )
		return colWidth;
	for( pnd = pnd.child( "w:gridCol" ); !pnd.empty(); pnd = pnd.next_sibling( "w:gridCol" ) ) {
		colWidth.push_back( pnd.attribute( "w:w" ).as_int() );
	}
	return colWidth;
}
int Table::GetColWidth( int idx ) const
{
	if( idx < 0 )
		return -1;
	pugi::xml_node pnd = m_nd.child( "w:tblGrid" );
	int i;
	for( i = 0, pnd = pnd.child( "w:gridCol" ); i < idx && !pnd.empty(); ++i, pnd = pnd.next_sibling( "w:gridCol" ) )
		;
	if( pnd.empty() )
		return -1;
	return pnd.attribute( "w:w" ).as_int();
}

int TCell::GetSpanNum() const
{
	int span_num = 0;
	pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:gridSpan" );
	if( !pnd.empty() )
		span_num = pnd.attribute( "w:val" ).as_int();
	return span_num;
}

VMergeTypeE TCell::GetVMergeType() const
{
	VMergeTypeE vmerge_type = VMergeTypeE::INVALID;
	pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:vMerge" );
	if( !pnd.empty() ) {
		pugi::xml_attribute attr = pnd.attribute( "w:val" );
		if( !attr.empty() && strcmp( attr.value(), "restart" ) == 0 )
			vmerge_type = VMergeTypeE::START;
		else
			vmerge_type = VMergeTypeE::CONT;
	}
	else {
		vmerge_type = VMergeTypeE::NONE;
	}
	return vmerge_type;
}

std::string TCell::GetText() const
{
	std::string text;
	bool is_first = true;
	// only get paragraph text, ignore table in table
	for( Element ele = GetFirstChild(); ele.IsValid(); ele = ele.GetNext() ) {
		if( ele.GetType() != ElementTypeE::PARAGRAPH )
			continue;
		Paragraph & par = static_cast<Paragraph &>( ele );
		if( is_first ) {
			is_first = false;
		}
		else {
			text.append( "\n" );
		}
		text.append( par.GetText() );
	}
	return text;
}

////////////////////////////////
END_NS_SPD
