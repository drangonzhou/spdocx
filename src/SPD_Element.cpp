// SPD_Element.cpp : spdocx docx element
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

#include "SPD_Element.h"
#include "SPD_Document.h"

#include <vector>
#include <string.h> // strcmp

BEGIN_NS_SPD
////////////////////////////////

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

pugi::xml_node Element::GetCreateChild( pugi::xml_node parent, const char * name )
{
	pugi::xml_node nd = parent.child( name );
	if( nd.empty() ) {
		nd = parent.append_child( name );
	}
	return nd;
}

pugi::xml_attribute Element::GetCreateAttr( pugi::xml_node nd, const char * name )
{
	pugi::xml_attribute attr = nd.attribute( name );
	if( attr.empty() ) {
		attr = nd.append_attribute( name );
	}
	return attr;
}

Element::Element( pugi::xml_node nd ) : m_nd( nd )
{
	m_type = Element::GetNodeType( nd );
}

ElementTypeE Element::GetNodeType( pugi::xml_node nd )
{
	ElementTypeE type = ElementTypeE::INVALID;
	if( !nd ) {
		type = ElementTypeE::INVALID;
	}
	else if( strcmp( nd.name(), "w:p" ) == 0 ) {
		type = ElementTypeE::PARAGRAPH;
	}
	else if( strcmp( nd.name(), "w:r" ) == 0 ) {
		if( !nd.child( "w:t" ).empty() || !nd.child( "w:drawing" ).empty() || !nd.child( "w:object" ).empty() ) {
			type = ElementTypeE::RUN;
		}
		// TODO (later) : other w:r, ex w:commentReference
		else {
			type = ElementTypeE::UNKNOWN;
		}
	}
	else if( strcmp( nd.name(), "w:hyperlink" ) == 0 ) {
		type = ElementTypeE::HYPERLINK;
	}
	else if( strcmp( nd.name(), "w:tbl" ) == 0 ) {
		type = ElementTypeE::TABLE;
	}
	else if( strcmp( nd.name(), "w:tr" ) == 0 ) {
		type = ElementTypeE::TABLE_ROW;
	}
	else if( strcmp( nd.name(), "w:tc" ) == 0 ) {
		type = ElementTypeE::TABLE_CELL;
	}
	// TODO (later) : other known element
	else {
		type = ElementTypeE::UNKNOWN;
	}
	return type;
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

Element Element::GetPrev( ElementTypeE type ) const
{
	if( m_type == ElementTypeE::INVALID )
		return Element();
	pugi::xml_node nd = m_nd.previous_sibling();
	while( !nd.empty() && GetNodeType( nd ) != type )
		nd = nd.previous_sibling();
	return Element( nd );
}

Element Element::GetNext( ElementTypeE type ) const
{
	if( m_type == ElementTypeE::INVALID )
		return Element();
	pugi::xml_node nd = m_nd.next_sibling();
	while( !nd.empty() && GetNodeType( nd ) != type )
		nd = nd.next_sibling();
	return Element( nd );
}

Element Element::GetChild( ElementTypeE type ) const
{
	if( m_type == ElementTypeE::INVALID || m_type == ElementTypeE::UNKNOWN )
		return Element();
	pugi::xml_node nd = m_nd.first_child();
	while( !nd.empty() && GetNodeType(nd) != type )
		nd = nd.next_sibling();
	return Element( nd );
}

int Element::DelChild( Element & child )
{
	// table del row, need special handle of cell with VMerge
	if( m_type == ElementTypeE::TABLE ) {
		return Table( *this ).DelRow( child );
	}
	// cell can not delete directly, use Table delete col, or use Cell span
	if( m_type == ElementTypeE::TABLE_ROW )
		return -1;

	if( child.m_nd.parent() != m_nd )
		return -1;
	bool ret = m_nd.remove_child( child.m_nd );
	return ret ? 0 : -1;
}

int Element::DelAllChild()
{
	if( m_type == ElementTypeE::TABLE_ROW )  // table row can not delete cell directly
		return -1;
	m_nd.remove_children();
	return 0;
}

const char * Paragraph::GetStyleId() const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	return attr.value();
}

const char * Paragraph::GetStyleName( const Document & doc ) const
{
	return doc.GetStyleName( GetStyleId() );
}

const char * Paragraph::GetNumId() const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:numPr" ).child( "w:numId" ).attribute( "w:val" );
	return attr.value();
}

int Paragraph::GetNumLevel() const
{
	pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:numPr" ).child( "w:ilvl" ).attribute( "w:val" );
	const char * val = attr.value();
	return ( val != nullptr ) ? atoi( val ) : -1;
}

int Paragraph::SetStyleId( const char * id )
{
	//pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:pStyle" ).attribute( "w:val" );
	pugi::xml_node nd = Element::GetCreateChild( m_nd, "w:pPr" );
	pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:pStyle" );
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

int Paragraph::SetStyleName( const char * name, const Document & doc )
{
	return SetStyleId( doc.GetStyleId( name ) );
}

int Paragraph::SetNumId( const char * id )
{
	//pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:numPr" ).child( "w:numId" ).attribute( "w:val" );
	pugi::xml_node nd = Element::GetCreateChild( m_nd, "w:pPr" );
	pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:numPr" );
	pugi::xml_node nd3 = nd2.child( "w:numId" );
	if( id == nullptr || id[0] == '\0' ) {
		if( nd3 ) {
			nd2.remove_child( "w:numId" );
		}
	}
	else {
		if( nd3.empty() ) {
			nd3 = nd2.append_child( "w:numId" );
		}
		Element::GetCreateAttr( nd3, "w:val" ).set_value( id );
	}
	return 0;
}

int Paragraph::SetNumLevel( int level )
{
	//pugi::xml_attribute attr = m_nd.child( "w:pPr" ).child( "w:numPr" ).child( "w:ilvl" ).attribute( "w:val" );
	pugi::xml_node nd = Element::GetCreateChild( m_nd, "w:pPr" );
	pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:numPr" );
	pugi::xml_node nd3 = nd2.child( "w:ilvl" );
	if( level == -1 ) {
		if( nd3 ) {
			nd2.remove_child( "w:ilvl" );
		}
	}
	else {
		if( nd3.empty() ) {
			nd3 = nd2.append_child( "w:ilvl" );
		}
		std::string lstr = std::to_string( level );
		Element::GetCreateAttr( nd3, "w:val" ).set_value( lstr.c_str() );
	}
	return 0;
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
	Table tbl = Element( nd );
	tbl.Reset();
	return tbl;
}

const char * Hyperlink::GetAnchor() const
{
	return m_nd.attribute( "w:anchor" ).value();
}

const char * Hyperlink::GetRelaId() const
{
	return m_nd.attribute( "r:id" ).value();
}

const Relationship * Hyperlink::GetRela( const Document & doc ) const
{
	return doc.GetRelationship( m_nd.attribute( "r:id" ).value() );
}

const std::string & Hyperlink::GetRelaType( const Document & doc ) const
{
	const Relationship * rela = doc.GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return std::string("");
	}
	return rela->m_type;
}

const std::string & Hyperlink::GetRelaTargetMode( const Document & doc ) const
{
	const Relationship * rela = doc.GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return std::string( "" );
	}
	return rela->m_targetMode;
}

const std::string & Hyperlink::GetRelaTarget( const Document & doc ) const
{
	const Relationship * rela = doc.GetRelationship( m_nd.attribute( "r:id" ).value() );
	if( rela == nullptr ) {
		return std::string( "" );
	}
	return rela->m_target;
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

int Hyperlink::SetRelaId( const char * id )
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

bool Run::IsText() const
{
	pugi::xml_node nd = m_nd.child( "w:t" );
	return nd.empty() ? false : true;
}

bool Run::IsPic() const
{
	// check xml node for picture
	pugi::xml_node nd = m_nd.child( "w:drawing" );
	return nd.empty() ? false : true;
}

bool Run::IsObject() const
{
	// check xml node for object
	pugi::xml_node nd = m_nd.child( "w:object" );
	return nd.empty() ? false : true;
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
		pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:color" );
		Element::GetCreateAttr( nd2, "w:val" ).set_value( color );
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
		pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:highlight" );
		Element::GetCreateAttr( nd2, "w:val" ).set_value( color );
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
		pugi::xml_node nd2 = Element::GetCreateChild( nd, "w:u" );
		Element::GetCreateAttr( nd2, "w:val" ).set_value( underline );
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
	m_nd.remove_child( "w:drawing" );
	m_nd.remove_child( "w:object" );
	Element::GetCreateChild( m_nd, "w:t" ).text().set( text );
	return;
}

const char * Run::GetPicRelaId() const
{
	return m_nd.child( "w:drawing" ).child( "wp:anchor" ).child( "a:graphic" ).child( "a:graphicData" )
		.child( "pic:pic" ).child( "pic:blipFill" ).child( "a:blip" ).attribute( "r:embed" ).value();
}

int  Run::GetPicData( const Document & doc, std::vector<char> & data ) const
{
	const char * picid = GetPicRelaId();
	if( picid == nullptr || picid[0] == '\0' )
		return -1;
	const Relationship * rela = doc.GetRelationship( picid );
	if( rela == nullptr )
		return -1;
	return doc.GetEmbedData( rela->m_target, data );
}

int Run::SetPic( const char * id )
{
	if( id == nullptr || id[0] == '\0' )
		return -1;
	m_nd.remove_child( "w:text" );
	m_nd.remove_child( "w:object" );
	pugi::xml_node nd = Element::GetCreateChild( m_nd, "w:drawing" );
	nd = Element::GetCreateChild( nd, "wp:anchor" );
	nd = Element::GetCreateChild( nd, "a:graphic" );
	nd = Element::GetCreateChild( nd, "a:graphicData" );
	nd = Element::GetCreateChild( nd, "pic:pic" );
	nd = Element::GetCreateChild( nd, "pic:blipFill" );
	nd = Element::GetCreateChild( nd, "a:blip" );
	Element::GetCreateAttr( nd, "r:embed" ).set_value( id );
	return 0;
}

int Run::SetPic( const char * id, Document & doc, const std::vector<char> & data )
{
	int ret = SetPic( id );
	if( ret < 0 )
		return ret;
	return doc.SetEmbedData( id, data );
}

int Run::SetPic( const char * id, Document & doc, std::vector<char> && data )
{
	int ret = SetPic( id );
	if( ret < 0 )
		return ret;
	return doc.SetEmbedData( id, std::move(data) );
}

const char * Run::GetObjectRelaId() const
{
	return m_nd.child( "w:object" ).child( "o:OLEObject" ).attribute( "r:id" ).value();
}

const char * Run::GetObjectProgId() const
{
	return m_nd.child( "w:object" ).child( "o:OLEObject" ).attribute( "ProgID" ).value();
}

const char * Run::GetObjectImgRelaId() const
{
	return m_nd.child( "w:object" ).child( "v:shape" ).child( "v:imagedata" ).attribute( "r:id" ).value();
}

int Run::GetObjectData( const Document & doc, std::vector<char> & data ) const
{
	const char * picid = GetObjectRelaId();
	if( picid == nullptr || picid[0] == '\0' )
		return -1;
	const Relationship * rela = doc.GetRelationship( picid );
	if( rela == nullptr )
		return -1;
	return doc.GetEmbedData( rela->m_target, data );
}

int Run::GetObjectImgData( const Document & doc, std::vector<char> & data ) const
{
	const char * picid = GetObjectImgRelaId();
	if( picid == nullptr || picid[0] == '\0' )
		return -1;
	const Relationship * rela = doc.GetRelationship( picid );
	if( rela == nullptr )
		return -1;
	return doc.GetEmbedData( rela->m_target, data );
}

int Run::SetObject( const char * objid, const char * progid, const char * imgid )
{
	if( objid == nullptr || objid[0] == '\0' || progid == nullptr || progid[0] == '\0' || imgid == nullptr || imgid[0] == '\0' )
		return -1;
	m_nd.remove_child( "w:text" );
	m_nd.remove_child( "w:drawing" );
	pugi::xml_node nd = Element::GetCreateChild( m_nd, "w:object" );
	pugi::xml_node nd2 = Element::GetCreateChild( nd, "o:OLEObject" );
	Element::GetCreateAttr( nd2, "r:id" ).set_value( objid );
	Element::GetCreateAttr( nd2, "ProgID" ).set_value( progid );
	nd2 = Element::GetCreateChild( nd, "v:shape" );
	nd2 = Element::GetCreateChild( nd2, "v:imagedata" );
	Element::GetCreateAttr( nd2, "r:id" ).set_value( imgid );
	return 0;
}

int Run::SetObject( const char * objid, const char * progid, const char * imgid, Document & doc, const std::vector<char> & objdata, const std::vector<char> & imgdata )
{
	int ret = SetObject( objid, progid, imgid );
	if( ret < 0 )
		return ret;
	ret = doc.SetEmbedData( objid, objdata );
	if( ret < 0 )
		return ret;
	return doc.SetEmbedData( imgid, imgdata );
}

int Run::SetObject( const char * objid, const char * progid, const char * imgid, Document & doc, std::vector<char> && objdata, std::vector<char> && imgdata )
{
	int ret = SetObject( objid, progid, imgid );
	if( ret < 0 )
		return ret;
	ret = doc.SetEmbedData( objid, std::move(objdata) );
	if( ret < 0 )
		return ret;
	return doc.SetEmbedData( imgid, std::move(imgdata) );
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
	for( auto & w : widths ) {
		if( w < 100 )
			return -1;
	}
	pugi::xml_node pnd = m_nd.child( "w:tblGrid" );
	int i;
	for( i = 0, pnd = pnd.child( "w:gridCol" ); i < num && !pnd.empty(); ++i, pnd = pnd.next_sibling( "w:gridCol" ) )
	{
		pnd.attribute( "w:w" ) = widths[i];
	}
	if( i != num ) {
		SPD_PR_DEBUG( "width num not match col num" );
	}

	return 0;
}

int Table::Reset( int row, int col )
{
	if( row < 1 || col < 1 || col > 80 )
		return -1;
	int i = 0;
	pugi::xml_node pnd;
	pnd = m_nd.child( "w:tblGrid" );
	if( pnd.empty() ) {
		pnd = m_nd.prepend_child( "w:tblGrid" );
	}
	pnd.remove_children();
	for( i = 0; i < col; ++i ) {
		pugi::xml_node cnd = pnd.append_child( "w:gridCol" );
		cnd.attribute( "w:w" ) = 8100 / col;
	}

	pnd = m_nd.child( "w:tblPr" );
	if( pnd.empty() ) {
		pnd = m_nd.prepend_child( "w:tblPr" );
	}
	
	m_nd.remove_child( "w:tr" );
	for( i = 0; i < row; ++i ) {
		pnd = m_nd.append_child( "w:tr" );
		pnd.append_child( "w:trPr" );
		int j = 0;
		for( j = 0; j < col; ++j ) {
			pugi::xml_node cnd = pnd.append_child( "w:tc" );
			cnd.append_child( "w:tcPr" );
		}
	}
	return 0;
}

TRow Table::AddChildTRow( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:tr" ) : m_nd.prepend_child( "w:tr" );
	return TRow( Element( nd ) );
}

int Table::DelRow( Element & row )
{
	if( row.m_nd.parent() != m_nd )
		return -1;

	TRow curr_row( row );
	Element child = curr_row.GetFirstChild();
	int curr_col = 0;
	while( child.IsValid() ) {
		TCell cell = child;
		int span = cell.GetSpanNum();
		VMergeTypeE vmt = cell.GetVMergeType();
		if( vmt == VMergeTypeE::START ) {
			TRow next_row1 = curr_row.GetNext();
			if( next_row1.IsValid() ) {
				int idx1 = 0;
				TCell next_cell1 = next_row1.GetCell( curr_col, idx1 );
				VMergeTypeE type = next_cell1.GetVMergeType();
				if( type == VMergeTypeE::CONT ) {
					TRow next_row2 = next_row1.GetNext();
					if( next_row2.IsValid() ) {
						int idx2 = 0;
						TCell next_cell2 = next_row2.GetCell( curr_col, idx2 );
						if( next_cell2.GetVMergeType() != VMergeTypeE::CONT ) {
							next_cell1.SetVMergeType( VMergeTypeE::NONE );
						}
						else {
							next_cell1.SetVMergeType( VMergeTypeE::START );
						}
					}
					else {
						next_cell1.SetVMergeType( VMergeTypeE::NONE );
					}
				}
				else {
					// should not happend
				}
			}
			else {
				// should not happend
			}
		}
		else if( vmt == VMergeTypeE::CONT ) {
			TRow prev_row1 = curr_row.GetPrev();
			if( prev_row1.IsValid() ) {
				int idx1 = 0;
				TCell prev_cell1 = prev_row1.GetCell( curr_col, idx1 );
				VMergeTypeE type = prev_cell1.GetVMergeType();
				if( type == VMergeTypeE::START ) {
					TRow next_row1 = curr_row.GetNext();
					if( next_row1.IsValid() ) {
						int idx2 = 0;
						TCell next_cell1 = next_row1.GetCell( curr_col, idx2 );
						if( next_cell1.GetVMergeType() != VMergeTypeE::CONT ) {
							prev_cell1.SetVMergeType( VMergeTypeE::NONE );
						}
						else {
							// no need to do anything
						}
					}
					else {
						prev_cell1.SetVMergeType( VMergeTypeE::NONE );
					}
				}
				else if( type == VMergeTypeE::CONT ) {
					// no need to do anything
				}
				else {
					// should not happend
				}
			}
			else {
				// should not happend
			}
			break;
		}
		child = child.GetNext();
		curr_col += span;
	}
	return m_nd.remove_child( row.m_nd );
}

Paragraph Table::AddSiblingParagraph( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:p", m_nd )
		: m_nd.parent().insert_child_before( "w:p", m_nd );
	return Paragraph( Element( nd ) );
}

Table Table::AddSiblingTable( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:tbl", m_nd )
		: m_nd.parent().insert_child_before( "w:tbl", m_nd );
	Table tbl = Element( nd );
	tbl.Reset();
	return tbl;
}

TCell TRow::GetCell( int col, int & idx )
{
	if( col < 0 ) {
		idx = -1;
		return TCell( Element() );
	}
	Element child = GetFirstChild();
	int curr_col = 0;
	int curr_idx = 0;
	while( child.IsValid() ) {
		TCell cell = TCell( child );
		int span = cell.GetSpanNum();
		if( curr_col <= col && col < curr_col + span ) {
			idx = curr_idx;
			return cell;
		}
		child = child.GetNext();
		curr_col += span;
		curr_idx += 1;
	}
	// not found
	idx = -1;
	return TCell( Element() );
}

TRow TRow::AddSiblingTRow( bool add_next )
{
	pugi::xml_node nd = add_next ? m_nd.parent().insert_child_after( "w:tr", m_nd )
		: m_nd.parent().insert_child_before( "w:tr", m_nd );
	// TODO : add cell
	return TRow( Element( nd ) );
}

int TCell::GetCol() const
{
	// TODO : get col
	int col = 0;
	Element ele = GetPrev();
	while( ele.GetType() == ElementTypeE::TABLE_CELL )
	{
		TCell cell = TCell( ele );
		col += cell.GetSpanNum();
		ele = ele.GetPrev();
	}
	return col;
}

int TCell::GetSpanNum() const
{
	int span_num = 1;
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

int TCell::GetVMergeNum() const
{
	VMergeTypeE vt = GetVMergeType();
	if( vt == VMergeTypeE::NONE ) {
		return 1;
	}
	else if( vt == VMergeTypeE::CONT ) {
		return -1;
	}
	else {
		int col = GetCol();
		int num = 1;
		TRow row = GetParent();
		TRow next_row = row.GetNext();
		while( next_row.IsValid() ) {
			int idx = 0;
			TCell next_cell = next_row.GetCell( col, idx );
			if( next_cell.GetVMergeType() == VMergeTypeE::CONT )
				num += 1;
			else
				break;
		}
		return num;
	}
	return -1; // never here
}

std::string TCell::GetText() const
{
	std::string text;
	bool is_first = true;
	// only get paragraph text, ignore table in table
	for( Element ele = GetFirstChild(); ele.IsValid(); ele = ele.GetNext() ) {
		if( ele.GetType() != ElementTypeE::PARAGRAPH )
			continue;
		Paragraph par = ele;
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
	if( num < 1 )
		return -1;
	if( GetVMergeType() == VMergeTypeE::CONT )
		return -1;
	int old_num = GetSpanNum();
	if( num == old_num ) {
		return 0;
	}
	else if( num < old_num ) {
		if( num == 1 ) {
			m_nd.child( "w:tcPr" ).remove_child( "w:gridSpan" );
		}
		else {
			pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:gridSpan" );
			pnd.attribute( "w:val" ) = num;
		}
		pugi::xml_node parent = m_nd.parent();
		pugi::xml_node last_nd = m_nd;
		for( int i = num; i < old_num; ++i ) {
			pugi::xml_node nnd = parent.insert_child_after( "w:tc", last_nd );
			nnd.append_child( "w:tcPr" );
			last_nd = nnd;
		}
		// TODO : if VMergeTypeE::START, also set following Vmerge Cell
	}
	else {
		// merge node to this
		int merge_num = num - old_num;
		std::vector< pugi::xml_node > merge_nds;
		merge_nds.resize( merge_num );
		pugi::xml_node last_nd = m_nd;
		int find_num = 0;
		while( 1 ) {
			pugi::xml_node nnd = last_nd.next_sibling();
			if( nnd.empty() )
				break;
			if( strcmp( nnd.name(), "w:tc" ) != 0 ) {
				last_nd = nnd;
				continue;
			}
			merge_nds[find_num++] = nnd;
			last_nd = nnd;
			if( find_num == merge_num )
				break;
		}
		if( find_num != merge_num )
			return -1;

		// set current node gridSpan, merge other cell to this
		pugi::xml_node pnd = m_nd.child( "w:tcPr" ).child( "w:gridSpan" );
		if( pnd.empty() ) {
			pnd = m_nd.child( "w:tcPr" ).append_child( "w:gridSpan" );
		}
		pnd.attribute( "w:val" ) = num;
		for( int i = 0; i < merge_num; ++i ) {
			pnd = merge_nds[i];
			pugi::xml_node cnd = pnd.first_child();
			while( !cnd.empty() ) {
				pugi::xml_node next = cnd.next_sibling();
				if( strcmp( cnd.name(), "w:p" ) == 0 || strcmp( cnd.name(), "w:tbl" ) == 0 )
					m_nd.append_move( cnd );
				cnd = next;
			}
			m_nd.parent().remove_child( pnd );
		}
		// TODO : if VMergeTypeE::START, also set following Vmerge Cell
	}
	return 0; 
}

int TCell::SetVMergeType( VMergeTypeE type )
{
	if( type == VMergeTypeE::NONE ) {
		pugi::xml_node pnd = m_nd.child( "w:tcPr" );
		if( !pnd.empty() ) {
			pnd.remove_child( "w:vMerge" );
		}
	}
	else if( type == VMergeTypeE::START ) {
		pugi::xml_node pnd = Element::GetCreateChild( m_nd, "w:tcPr" );
		pnd = Element::GetCreateChild( pnd, "w:vMerge" );
		Element::GetCreateAttr( pnd, "w:val" ).set_value( "restart" );
	}
	else if( type == VMergeTypeE::CONT ) {
		pugi::xml_node pnd = Element::GetCreateChild( m_nd, "w:tcPr" );
		pnd = Element::GetCreateChild( pnd, "w:vMerge" );
		Element::GetCreateAttr( pnd, "w:val" ).set_value( "continue" );
	}
	else {
		return -1;
	}
	return 0;
}

int TCell::SetVMergeNum( int num )
{
	// TODO : need to check following row enough or not
	// TODO : need to deal with span in following row
	return -1;
}

Paragraph TCell::AddChildParagraph( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:p" ) : m_nd.prepend_child( "w:p" );
	return Paragraph( Element( nd ) );
}

Table TCell::AddChildTable( bool add_back )
{
	pugi::xml_node nd = add_back ? m_nd.append_child( "w:tbl" ) : m_nd.prepend_child( "w:tbl" );
	Table tbl = Element( nd );
	tbl.Reset();
	return tbl;
}

////////////////////////////////
END_NS_SPD
