// SPD_Document.cpp : spdocx document
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

#include "SPD_Document.h"
#include "SPD_NewDoc.h"

#include <zip.h>
#include <memory>
#include <string.h> // strcmp

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
	Close();

	m_files["[Content_Types].xml"].assign( s_content_type, s_content_type + strlen( s_content_type ) + 1 );
	m_files["docProps/app.xml"].assign( s_docprops_app, s_docprops_app + strlen( s_docprops_app ) + 1 );
	m_files["docProps/core.xml"].assign( s_docprops_core, s_docprops_core + strlen( s_docprops_core ) + 1 );
	m_files["_rels/.rels"].assign( s_rels_rels, s_rels_rels + strlen( s_rels_rels ) + 1 );
	m_files["word/document.xml"].assign( s_word_document, s_word_document + strlen( s_word_document ) + 1 );
	m_files["word/styles.xml"].assign( s_word_styles, s_word_styles + strlen( s_word_styles ) + 1 );
	m_files["word/numbering.xml"].assign( s_word_numbering, s_word_numbering + strlen( s_word_numbering ) + 1 );
	m_files["word/fontTable.xml"].assign( s_word_font_table, s_word_font_table + strlen( s_word_font_table ) + 1 );
	m_files["word/settings.xml"].assign( s_word_settings, s_word_settings + strlen( s_word_settings ) + 1 );
	m_files["word/webSettings.xml"].assign( s_word_web_settings, s_word_web_settings + strlen( s_word_web_settings ) + 1 );
	m_files["word/_rels/document.xml.rels"].assign( s_word_rels, s_word_rels + strlen( s_word_rels ) + 1 );
	m_files["word/theme/theme1.xml"].assign( s_word_theme, s_word_theme + strlen( s_word_theme ) + 1 );

	int ret = SPD_ERR_ERROR;
	if( ( ret = parse_files() ) != SPD_ERR_OK ) {
		Close();
	}
	return SPD_ERR_OK;
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
		if( rdsize != (int64_t)zstat.size ) {
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

	int ret = SPD_ERR_ERROR;
	if( ( ret = read_zip( fname, m_files ) ) < 0 )
		return ret;
	if( ( ret = parse_files()) != SPD_ERR_OK ) {
		Close();
	}
	else {
		m_fname = fname;
	}
	return ret;
}

int Document::parse_files()
{
	int ret = SPD_ERR_ERROR;
	if( (ret = read_xml( "word/document.xml", &m_doc )) < 0 ) {
		SPD_PR_INFO( "read_zip_xml() [word/document.xml] failed, err %d", ret );
		return ret;
	}

	// simple check
	pugi::xml_node nd = m_doc.document_element();
	if( strcmp( nd.name(), "w:document" ) != 0 || strcmp( nd.first_child().name(), "w:body" ) != 0 ) {
		SPD_PR_INFO( "bad xml content" );
		return SPD_ERR_OPEN_XML;
	}

	if (load_style() < 0) {
		SPD_PR_INFO("bad style content");
		return SPD_ERR_OPEN_XML;
	}

	load_rela();  // ignore error

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
	int ret = SPD_ERR_ERROR;
	if( !fname.empty() ) {
		m_fname = fname;
	}
	if( m_fname.empty() ) {
		SPD_PR_INFO( "error empty zip filename" );
		return SPD_ERR_SAVE_ZIP;
	}

	write_xml( "word/document.xml", m_doc );
	
	write_style();
	write_rela();

	// save zip file
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
			pugi::xml_node ppr = nd.child( "w:pPr" );
			pugi::xml_node numpr = ppr.child( "w:numPr" );
			if( numpr ) {
				pugi::xml_node numid = numpr.child( "w:numId" );
				if( numid ) {
					style.m_numId = numid.attribute( "w:val" ).value();
				}
				else {
					style.m_numId = "";
				}
				pugi::xml_node numlevel = numpr.child( "w:ilvl" );
				if( numlevel ) {
					const char * val = numlevel.attribute( "w:val" ).value();
					style.m_numLevel = (val != nullptr) ? atoi( val ) : 0;
				}
				else {
					style.m_numLevel = 0;
				}
			}
			else {
				style.m_numId = "";
				style.m_numLevel = 0;
			}
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
			// Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
			size_t pos = rela.m_type.find_last_of( '/' );
			if( pos != std::string::npos ) {
				rela.m_type = rela.m_type.substr( pos + 1 );
			}
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

const StyleLite * Document::GetStyle( const char* id ) const
{
	if( id == nullptr || id[0] == '\0' )
		return nullptr;
	auto it = m_style.find( std::string( id ) );
	if( it == m_style.end() )
		return nullptr;
	return &it->second;
}

std::vector< const StyleLite*> Document::GetAllStyle() const
{
	std::vector< const StyleLite* > ret;
	for( auto it = m_style.begin(); it != m_style.end(); ++it ) {
		ret.push_back( &(it->second) );
	}
	return ret;
}

int Document::AddStyle( const StyleLite & style )
{
	if( style.m_id.empty() )
		return -1;
	m_style[style.m_id] = style;
	return 0;
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

std::vector< const Relationship* > Document::GetAllRelationship() const
{
	std::vector< const Relationship * > ret;
	for( auto it = m_rela.begin(); it != m_rela.end(); ++it ) {
		ret.push_back( &(it->second) );
	}
	return ret;
}
int Document::AddRelationship( const Relationship& rela )
{
	if( rela.m_id.empty() )
		return -1;
	m_rela[rela.m_id] = rela;
	return 0;
}

int Document::write_style()
{
	pugi::xml_document doc;
	int ret = read_xml( "word/styles.xml", &doc );
	if( ret < 0 ) {
		m_files["word/styles.xml"].assign( s_word_styles, s_word_styles + strlen( s_word_styles ) + 1 );
		ret = read_xml( "word/styles.xml", &doc );
	}

	pugi::xml_node pnd = doc.document_element().first_child(); // child( "w:styles" );
	for( auto & it : m_style )
	{
		pugi::xml_node nd = pnd.find_child( [&it]( pugi::xml_node cnd ) { return cnd.attribute( "w:styleId" ).value() == it.first; } );
		if( !nd ) {
			nd = pnd.append_child( "w:style" );
		}
		Element::GetCreateAttr( nd, "w:type" ).set_value( it.second.m_type.c_str() );
		pugi::xml_node cnd = Element::GetCreateChild( nd, "w:name" );
		Element::GetCreateAttr( nd, "w:val" ).set_value( it.second.m_name.c_str() );
		if( !it.second.m_numId.empty() ) {
			cnd = Element::GetCreateChild( nd, "w:pPr" );
			pugi::xml_node cnd2 = Element::GetCreateChild( cnd, "w:numPr" );
			pugi::xml_node cnd3 = Element::GetCreateChild( cnd2, "w:numId" );
			Element::GetCreateAttr( cnd3, "w:val" ).set_value( it.second.m_numId.c_str() );
			cnd3 = Element::GetCreateChild( cnd2, "w:ilvl" );
			Element::GetCreateAttr( cnd3, "w:numId" ).set_value( it.second.m_numId.c_str() );
		}
	}
	return 0;
}

int Document::write_rela()
{
	pugi::xml_document doc;
	int ret = read_xml( "word/_rels/document.xml.rels", &doc );
	if( ret < 0 ) {
		m_files["word/_rels/document.xml.rels"].assign( s_word_rels, s_word_rels + strlen( s_word_rels ) + 1 );
		ret = read_xml( "word/_rels/document.xml.rels", &doc );
	}
	pugi::xml_node pnd = doc.document_element().first_child(); // child( "Relationships" );
	for( auto & it : m_rela )
	{
		pugi::xml_node nd = pnd.find_child( [&it]( pugi::xml_node cnd ) { return cnd.attribute( "Id" ).value() == it.first; } );
		if( !nd ) {
			nd = pnd.append_child( "Relationship" );
		}
		// Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
		std::string type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
		type += it.second.m_type;
		Element::GetCreateAttr( nd, "Type" ).set_value( type.c_str() );
		Element::GetCreateAttr( nd, "Target" ).set_value( it.second.m_target.c_str() );
		if( !it.second.m_targetMode.empty() ) {
			Element::GetCreateAttr( nd, "TargetMode" ).set_value( it.second.m_targetMode.c_str() );
		}
	}
	return 0;
}

////////////////////////////////
END_NS_SPD
