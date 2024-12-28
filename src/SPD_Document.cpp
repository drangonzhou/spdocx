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
#include "zip.h"

#include <memory>
#include <cstring>

BEGIN_NS_SPD
////////////////////////////////

Document::Document()
{
	 
}

Document::~Document()
{
	Close();
}

int Document::New()
{
	int ret = 0;
	Close();
	// TODO : set default style
	// TODO : need impl
	return -1;
}

int Document::read_zip( const std::string & fname, std::map< std::string, std::vector<char> > & files )
{
	zip_t * zip = nullptr;
	int ret = 0;
	zip = zip_open( fname.c_str(), 0, &ret );
	if( zip == nullptr ) {
		SPD_PR_INFO( "zip_open() [%s] failed, err %d", fname.c_str(), ret );
		return SPD_ERR_OPEN_ZIP;
	}
	//std::unique_ptr<zip_t, void( * )( zip_t * )> guard1( zip, []( zip_t * z ) { if( z != nullptr ) zip_close( z ); } );

	zip_int64_t num = zip_get_num_entries( zip, 0 );
	for( zip_int64_t i = 0; i < num; ++i ) {
		zip_stat_t zstat = { 0 };
		if( ( ret = zip_stat_index( zip, i, 0, &zstat ) ) != 0 ) {
			SPD_PR_DEBUG( "zip_stat() [%d] failed, err %d", (int)i, ret );
			continue; // ignore error
		}
		if( zstat.name == nullptr || zstat.size == 0 ) {
			SPD_PR_DEBUG( "zip_stat() [%d] empty, skip", (int)i );
			continue;
		}
		zip_file_t * zfile = zip_fopen_index( zip, i, 0 );
		if( ! zfile ) {
			SPD_PR_DEBUG( "zip_fopen_index() [%d] failed, skip", (int)i );
			continue;
		}
		//std::unique_ptr<zip_file_t, void( * )( zip_file_t * )> guard2( zfile, []( zip_file_t * z ) { if( z != nullptr ) zip_fclose( z ); } );
		std::vector<char> & fbuf = files[zstat.name];
		fbuf.resize( zstat.size + 1 );
		zip_int64_t rdsize = zip_fread( zfile, &fbuf[0], zstat.size );
		if( rdsize != zstat.size ) {
			SPD_PR_DEBUG( "zip_stat() [%d] read file error, fsize %d, read %d", (int)zstat.size, (int)rdsize );
		}
		fbuf[rdsize] = '\0';
		zip_fclose( zfile ), zfile = nullptr;
	}
	zip_close( zip ), zip = nullptr;
	return SPD_ERR_OK;
}

int Document::write_zip( const std::string & fname, const std::map< std::string, std::vector<char> > & files )
{
	zip_t * zip = nullptr;
	int ret = 0;
	zip = zip_open( fname.c_str(), ZIP_CREATE | ZIP_TRUNCATE, &ret );
	if( zip == nullptr ) {
		SPD_PR_INFO( "zip_open() [%s] failed, err %d", fname.c_str(), ret );
		return SPD_ERR_OPEN_ZIP;
	}

	for( const auto & it : files ) {
		const std::vector<char> & fbuf = it.second;
		if( fbuf.empty() ) {
			continue; // skip empty file, every file should has an extra '\0'
		}
		zip_source_t * src = zip_source_buffer( zip, &fbuf[0], fbuf.size() - 1, 0 );
		if( !src ) {
			SPD_PR_INFO( "zip_source_buffer() failed" );
			continue;
		}
		if( zip_file_add( zip, it.first.c_str(), src, ZIP_FL_ENC_GUESS ) < 0 ) {
			SPD_PR_INFO( "zip_file_add() failed" );
			zip_source_free( src ); // Free the source if adding fails
			continue;
		}
		// no need to free src
	}

	if( zip_close( zip ) != 0 ) {
		SPD_PR_INFO( "zip_close() failed" );
		return SPD_ERR_SAVE_ZIP;
	}
	return SPD_ERR_OK;
}

int Document::Open( const std::string & fname )
{
	Close();

	int ret = -1;
	if( ( ret = read_zip( fname, m_files ) ) < 0 )
		return ret;

	if( (ret = read_xml( "word/document.xml", &m_doc )) < 0 ) {
		SPD_PR_INFO( "read_zip_xml() [word/document.xml] failed, err %d", ret );
		Close();
		return ret;
	}

	// simple check
	pugi::xml_node nd = m_doc.document_element();
	if( strcmp( nd.name(), "w:document" ) != 0 || strcmp( nd.first_child().name(), "w:body" ) != 0 ) {
		SPD_PR_INFO( "bad xml content" );
		Close();
		return SPD_ERR_OPEN_XML;
	}

	if (load_style() < 0) {
		SPD_PR_INFO("bad style content");
		Close();
		return SPD_ERR_OPEN_XML;
	}

	load_rela();  // ignore error

	m_fname = fname;

	return SPD_ERR_OK;
}

class BufferWriter : public pugi::xml_writer 
{
public:
	std::vector<char> & m_buf;
	BufferWriter( std::vector<char> & buf ) : m_buf( buf )
	{
		m_buf.resize( 0 );
		m_buf.reserve( 64 * 1024 );
	}
	virtual void write( const void * data, size_t size ) 
	{
		m_buf.insert( m_buf.end(), static_cast<const char*>(data), static_cast<const char*>(data) + size );
	}
};

int Document::write_xml( const std::string & fname, const pugi::xml_document & doc )
{
	std::vector<char> & fbuf = m_files[fname];
	// convert xml to zip file
	BufferWriter writer(fbuf);
	m_doc.save( writer );
	fbuf.push_back( '\0' );

	return SPD_ERR_OK;
}

int Document::Save( const std::string & fname )
{
	int ret = -1;

	write_xml( "word/document.xml", m_doc );
	
	// TODO : style and rela

	// save zip file
	if( !fname.empty() )
		m_fname = fname;
	if( m_fname.empty() ) {
		SPD_PR_INFO( "error empty zip filename" );
		return SPD_ERR_SAVE_ZIP;
	}
	if( ( ret = write_zip( m_fname, m_files ) ) != SPD_ERR_OK ) {
		return ret;
	}
	return SPD_ERR_OK;
}

int Document::Close()
{
	m_fname.clear();
	m_files.clear();
	m_isModified = false;

	m_doc.reset();
	m_style.clear();
	m_rela.clear();
	return 0;
}

int Document::load_style()
{
	pugi::xml_document doc;
	int ret = read_xml( "word/styles.xml", &doc );
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
	int ret = read_xml( "word/_rels/document.xml.rels", &doc );
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

int Document::read_xml( const std::string & fname, pugi::xml_document * doc ) const
{
	std::map< std::string, std::vector<char> >::const_iterator it = m_files.find( fname );
	if( it == m_files.end() )
		return SPD_ERR_OPEN_XML;

	const std::vector<char> fbuf = it->second;
	pugi::xml_parse_result xmlret = doc->load_buffer( &fbuf[0], (size_t)fbuf.size() );
	if( !xmlret ) {
		SPD_PR_DEBUG( "load xml [%s] failed, err : %s", fname.c_str(), xmlret.description() );
		return SPD_ERR_OPEN_XML;
	}

	return SPD_ERR_OK;
}

Element Document::GetFirstElement() const
{
	// <w:document><w:body><w:p>...</w:p>...</w:body></w:document>
	pugi::xml_node nd = m_doc.document_element().first_child().first_child();
	return Element( nd );
}

Paragraph Document::AddChildParagraph( bool add_back )
{
	pugi::xml_node parent = m_doc.document_element().first_child();
	pugi::xml_node nd = add_back ? parent.append_child( "w:p" ) : parent.prepend_child( "w:p" );
	return Paragraph( Element( nd ) );
}

Table Document::AddChildTable( bool add_back )
{
	pugi::xml_node parent = m_doc.document_element().first_child();
	pugi::xml_node nd = add_back ? parent.append_child( "w:tbl" ) : parent.prepend_child( "w:tbl" );
	Table tbl = Element( nd );
	tbl.Reset();
	return tbl;
}

int Document::DelChild( Element & child )
{
	pugi::xml_node parent = m_doc.document_element().first_child();
	bool ret = parent.remove_child( child.m_nd );
	return ret ? 0 : -1;
}

int Document::DelAllChild()
{
	pugi::xml_node parent = m_doc.document_element().first_child();
	parent.remove_children();
	return 0;
}

const char * Document::GetStyleName( const char * id ) const
{
	if( id == nullptr || id[0] == '\0' )
		return "";
	auto it = m_style.find( std::string( id ) );
	if( it == m_style.end() )
		return "";
	return it->second.m_name.c_str();
}

const char * Document::GetStyleId( const char * name ) const
{
	if( name == nullptr || name[0] == '\0' )
		return "";
	auto it = m_style.begin();
	for( ; it != m_style.end(); ++it ) {
		if( it->second.m_name == name )
			return it->second.m_id.c_str();
	}
	return "";
}

const Relationship * Document::GetRelationship( const char * id ) const
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
