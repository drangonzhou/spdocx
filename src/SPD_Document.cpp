// SPD_Document.cpp : spdocx document
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

#include "SPD_Document.h"

#include <string.h>

BEGIN_NS_SPD
////////////////////////////////

Document::Document()
	: m_zip( nullptr )
{
	 
}

Document::~Document()
{
	Close();
}

int Document::Open( const char * fname )
{
	if( m_zip != nullptr )
		Close();

	int ret = 0;
	m_zip = zip_open( fname, 0, &ret );
	if( m_zip == nullptr ) {
		SPD_PR_INFO( "zip_open() [%s] failed, err %d", fname, ret );
		return SPD_ERR_OPEN_ZIP;
	}

	ret = read_zip_xml( "word/document.xml", &m_doc );
	if( ret < 0 ) {
		SPD_PR_INFO( "read_zip_xml() [word/document.xml] failed, err %d", ret );
		Close();
		return ret;
	}

	// simple check
	pugi::xml_node nd = m_doc.document_element();
	if( strcmp( nd.name(), "w:document" ) != 0 || strcmp( nd.first_child().name(), "w:body" ) != 0 || load_style() < 0 ) {
		SPD_PR_INFO( "bad xml content" );
		Close();
		return SPD_ERR_OPEN_XML;
	}

	load_rela();  // ignore error

	return 0;
}

int Document::Save( const char * fname )
{
	// TODO : need imp

	return -1;
}

int Document::Close()
{
	if( m_zip != nullptr ) {
		zip_discard( m_zip ), m_zip = nullptr;
	}
	m_doc.reset();
	m_style.clear();
	m_rela.clear();
	return 0;
}

int Document::load_style()
{
	pugi::xml_document doc;
	int ret = read_zip_xml( "word/styles.xml", &doc );
	if( ret < 0 )
		return ret;

	pugi::xml_node rootnd = doc.document_element();
	if( strcmp( rootnd.name(), "w:styles" ) != 0 ) {
		SPD_PR_DEBUG( "document_element() is not w:styles" );
		return ret;
	}
	for( pugi::xml_node nd = rootnd.first_child(); nd; nd = nd.next_sibling() )	{
		if( strcmp( nd.name(), "w:style" ) == 0 ) {
			std::string id = nd.attribute( "w:styleId" ).value();
			if( id.empty() ) {
				SPD_PR_DEBUG( "empty styleId" );
				continue;
			}
			StyleLite & style = m_style[id];
			style.m_id = id;
			style.m_type = nd.attribute( "w:type" ).value();
			style.m_name = nd.child( "w:name" ).attribute( "w:val" ).value();
		}
		else if( strcmp( nd.name(), "w:docDefaults" ) == 0 || strcmp( nd.name(), "w:latentStyles" ) == 0 ) {
			// Ignore
			continue;
		}
		else {
			// unknown tag
			SPD_PR_DEBUG( "unknown tag [%s]", nd.name() );
			continue;
		}
	}
	return SPD_ERR_OK;
}

int Document::load_rela()
{
	pugi::xml_document doc;
	int ret = read_zip_xml( "word/_rels/document.xml.rels", &doc );
	if( ret < 0 )
		return ret;

	pugi::xml_node rootnd = doc.document_element();
	if( strcmp( rootnd.name(), "Relationships" ) != 0 ) {
		SPD_PR_DEBUG( "document_element() is not Relationships" );
		return ret;
	}
	for( pugi::xml_node nd = rootnd.first_child(); nd; nd = nd.next_sibling() ) {
		if( strcmp( nd.name(), "Relationship" ) == 0 ) {
			std::string id = nd.attribute( "Id" ).value();
			if( id.empty() ) {
				SPD_PR_DEBUG( "empty styleId" );
				continue;
			}
			Relationship & rela = m_rela[id];
			rela.m_id = id;
			rela.m_type = nd.attribute( "Type" ).value();
			rela.m_target = nd.attribute( "Target" ).value();
			rela.m_targetMode = nd.attribute( "TargetMode" ).value();
		}
		else {
			// unknown tag
			SPD_PR_DEBUG( "unknown tag [%s]", nd.name() );
			continue;
		}
	}
	return SPD_ERR_OK;

}

int Document::read_zip_xml( const char * zip_fname, pugi::xml_document * doc )
{
	int ret;
	zip_stat_t zst;
	ret = zip_stat( m_zip, zip_fname, 0, &zst );
	if( ret != 0 || zst.size == 0 ) {
		SPD_PR_DEBUG( "zip_stat() [%s] failed, err : %s", zip_fname, zip_strerror( m_zip ) );
		return SPD_ERR_OPEN_ZIP;
	}

	zip_file_t * zf = zip_fopen( m_zip, zip_fname, 0 );
	if( zf == nullptr ) {
		SPD_PR_DEBUG( "zip_fopen() [%s] failed, err : %s", zip_fname, zip_strerror( m_zip ) );
		return SPD_ERR_OPEN_ZIP;
	}
	char * buf = new char[zst.size + 1];
	zip_int64_t size = (int)zip_fread( zf, (void *)buf, zst.size );
	buf[zst.size] = '\0';
	if( size != zst.size ) {
		SPD_PR_DEBUG( "zip_fread() [%s] failed, err : %s", zip_fname, zip_strerror( m_zip ) );
		delete[] buf;
		zip_fclose( zf );
		return SPD_ERR_OPEN_XML;
	}
	zip_fclose( zf ), zf = nullptr;

	pugi::xml_parse_result xmlret = doc->load_buffer( buf, (size_t)zst.size );
	delete[] buf, buf = nullptr;
	if( !xmlret ) {
		SPD_PR_DEBUG( "load xml [%s] failed, err : %s", zip_fname, xmlret.description() );
		return SPD_ERR_OPEN_XML;
	}

	return SPD_ERR_OK;
}

RefPtr<Element> Document::GetFirstElement()
{
	// <w:document><w:body><w:p>...</w:p>...</w:body></w:document>
	pugi::xml_node nd = m_doc.document_element().first_child().first_child();
	return Element::CreateElement( this, nd );
}

const char * Document::get_style_name( const char * id )
{
	if( id == nullptr || id[0] == '\0' )
		return "";
	auto it = m_style.find( std::string( id ) );
	if( it == m_style.end() )
		return "";
	return it->second.m_name.c_str();
}

const Relationship * Document::get_relationship( const char * id )
{
	if( id == nullptr || id[0] == '\0' )
		return nullptr;
	auto it = m_rela.find( std::string( id ) );
	if( it == m_rela.end() )
		return nullptr;
	return &(it->second);
}

////////////////////////////////
END_NS_SPD
