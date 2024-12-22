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

int Element::DelChild( Element & child )
{
	bool ret = m_nd.remove_child( child.m_nd );
	return ret ? 0 : -1;
}

int Element::DelAllChild()
{
	m_nd.remove_children();
	return 0;
}

const char * Paragraph::GetStyleId() const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	return attr.value();
}

const char * Paragraph::GetStyleName( const Document * doc ) const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	return doc->GetStyleName( attr.value() );
}

int Paragraph::SetStyleId( const char * id )
{
	//pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	pugi::xml_node nd = m_nd.child( "w:pPr" );
	if( nd.empty() ) {
		nd = m_nd.append_child( "w:pPr" );
	}
	pugi::xml_node nd2 = nd.child( "w:pStyle" );
	if( nd2.empty() ) {
		nd2 = nd.append_child( "w:pStyle" );
	}
	pugi::xml_attribute attr = nd2.attribute( "w:val" );
	if( id == nullptr || id[0] == '\0' ) {
		if( !attr.empty() )
			nd2.remove_attribute( attr );
	}
	else {
		if( attr.empty() ) {
			attr = nd2.append_attribute( "w:val" );
		}
		attr.set_value( id );
	}
	return 0;
}

int Paragraph::SetStyleName( const char * name, const Document * doc )
{
	const char * id = doc->GetStyleId( name );
	return SetStyleId( id );
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

Run Paragraph::AddChildRun( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:r" ) : m_nd.prepend_child( "w:r" );
	return Run( Element( nd ) );
}

Hyperlink Paragraph::AddChildHyperlink( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:hyperlink" ) : m_nd.prepend_child( "w:hyperlink" );
	return Hyperlink( Element( nd ) );
}

Paragraph Paragraph::AddSiblingParagraph( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:p", m_nd )
		: m_nd.parent().insert_child_before( "w:p", m_nd );
	return Paragraph( Element( nd ) );
}

Table Paragraph::AddSiblingTable( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:tbl", m_nd ) 
		: m_nd.parent().insert_child_before( "w:tbl", m_nd );
	return Table( Element( nd ) );
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

int Hyperlink::SetAnchor( const char * anchor )
{
	pugi::xml_attribute attr = m_nd.attribute( "w:anchor" );
	if( anchor == nullptr || anchor[0] == '\0' ) {
		if( !attr.empty() )
			m_nd.remove_attribute( attr );
	}
	else {
		if( attr.empty() )
			attr = m_nd.append_attribute( "w:anchor" );
		attr.set_value( anchor );
	}
	return 0;
}

int Hyperlink::SetRelationshipId( const char * id )
{
	pugi::xml_attribute attr = m_nd.attribute( "r:id" );
	if( id == nullptr || id[0] == '\0' ) {
		if( !attr.empty() )
			m_nd.remove_attribute( attr );
	}
	else {
		if( attr.empty() )
			attr = m_nd.append_attribute( "r:id" );
		attr.set_value( id );
	}
	return 0;
}

Run Hyperlink::AddChildRun( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:r" ) : m_nd.prepend_child( "w:r" );
	return Run( Element( nd ) );
}

Hyperlink Hyperlink::AddSiblingHyperlink( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:hyperlink", m_nd )
		: m_nd.parent().insert_child_before( "w:hyperlink", m_nd );
	return Hyperlink( Element( nd ) );
}

Run Hyperlink::AddSiblingRun( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:r", m_nd )
		: m_nd.parent().insert_child_before( "w:r", m_nd );
	return Run( Element( nd ) );
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

int Run::SetColor( const char * color )
{
	//pugi::xml_attribute attr = m_nd.child( "w:rPr" ).child( "w:color" ).attribute( "w:val" );
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( color == nullptr || color[0] == '\0' ) {
		if( !nd.empty() ) {
			nd.remove_child( "w:color" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	else {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:color" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:color" );
		pugi::xml_attribute attr = nd2.attribute( "w:val" );
		if( attr.empty() )
			attr = nd2.append_attribute( "w:val" );
		attr.set_value( color );
	}
	return 0;
}

int Run::SetHighline( const char * color )
{
	//pugi::xml_attribute attr = m_nd.child( "w:rPr" ).child( "w:highlight" ).attribute( "w:val" );
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( color == nullptr || color[0] == '\0' ) {
		if( !nd.empty() ) {
			nd.remove_child( "w:highlight" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	else {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:highlight" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:highlight" );
		pugi::xml_attribute attr = nd2.attribute( "w:val" );
		if( attr.empty() )
			attr = nd2.append_attribute( "w:val" );
		attr.set_value( color );
	}
	return 0;
}

int Run::SetBold( bool bold )
{
	// !m_nd.child( "w:rPr" ).child( "w:b" ).empty();
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( bold ) {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:b" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:b" );
	}
	else {
		if( !nd.empty() ) {
			nd.remove_child( "w:b" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	return 0;
}

int Run::SetItalic( bool italic )
{
	// !m_nd.child( "w:rPr" ).child( "w:i" ).empty();
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( italic ) {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:i" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:i" );
	}
	else {
		if( !nd.empty() ) {
			nd.remove_child( "w:i" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	return 0;
}

int Run::SetUnderline( const char * underline )
{
	// m_nd.child( "w:rPr" ).child( "w:u" ).attribute( "w:val" ).value();
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( underline == nullptr || underline[0] == '\0' ) {
		if( !nd.empty() ) {
			nd.remove_child( "w:u" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	else {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:u" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:u" );
		pugi::xml_attribute attr = nd2.attribute( "w:val" );
		if( attr.empty() )
			attr = nd2.append_attribute( "w:val" );
		attr.set_value( underline );
	}
	return 0;
}

int Run::SetStrike( bool strike )
{
	// m_nd.child( "w:rPr" ).child( "w:strike" ).empty();
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( strike ) {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:strike" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:strike" );
		nd.remove_child( "w:dstrike" );
	}
	else {
		if( !nd.empty() ) {
			nd.remove_child( "w:strike" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	return 0;
}

int Run::SetDoubleStrike( bool dstrike )
{
	// m_nd.child( "w:rPr" ).child( "w:dstrike" ).empty();
	pugi::xml_node nd = m_nd.child( "w:rPr" );
	if( dstrike ) {
		if( nd.empty() )
			nd = m_nd.append_child( "w:rPr" );
		pugi::xml_node nd2 = nd.child( "w:dstrike" );
		if( nd2.empty() )
			nd2 = nd.append_child( "w:dstrike" );
		nd.remove_child( "w:strike" );
	}
	else {
		if( !nd.empty() ) {
			nd.remove_child( "w:dstrike" );
			if( nd.first_child().empty() )
				m_nd.remove_child( nd );
		}
	}
	return 0;
}

void Run::SetText( const char * text )
{
	m_nd.child( "w:t" ).text().set( text );
	return;
}

Hyperlink Run::AddSiblingHyperlink( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:hyperlink", m_nd )
		: m_nd.parent().insert_child_before( "w:hyperlink", m_nd );
	return Hyperlink( Element( nd ) );
}

Run Run::AddSiblingRun( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:r", m_nd )
		: m_nd.parent().insert_child_before( "w:r", m_nd );
	return Run( Element( nd ) );
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

int Table::AddCol( int index, int num )
{
	// TODO : update column size

	// add cell to every row
	return -1;
}

int Table::DelCol( int index, int num )
{
	// TODO : only one column can not delete
	
	// update column

	// del cell in every row
	return -1;
}

int Table::SetColWidth( const std::vector<int> widths )
{
	int num = GetColNum();
	if( (int)widths.size() != num )
		return -1;
	// TODO : 
	return -1;
}

int Table::AddRow( int index, int num )
{
	// TODO 
	return -1;
}

int Table::DelRow( int index, int num )
{
	// TODO 
	return -1;
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

int TCell::SetSpanNum( int num )
{
	// TODO :
	return -1;
}

int TCell::SetVmergeNum( int num )
{
	// TODO : 
	return -1;
}

////////////////////////////////
END_NS_SPD
