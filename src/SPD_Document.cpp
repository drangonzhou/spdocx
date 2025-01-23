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

#include <zip.h>
#include <memory>
#include <string.h> // strcmp

BEGIN_NS_SPD
////////////////////////////////

static const char * s_content_type = R"(<?xml encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" />
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" />
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" />
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml" />
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" />
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
</Types>
)";

static const char * s_docprops_app = R"(<?xml encoding="UTF-8" standalone="yes"?>
<Properties 
    xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" 
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Template>Normal.dotm</Template>
    <TotalTime>0</TotalTime>
    <Application>spdocx</Application>
    <DocSecurity>0</DocSecurity>
</Properties>
)";

static const char * s_docprops_core = R"(<?xml encoding="UTF-8" standalone="yes"?>
<cp:coreProperties 
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    <dc:title></dc:title>
    <dc:subject></dc:subject>
    <dc:creator>spdocx</dc:creator>
    <dc:description></dc:description>
    <cp:keywords></cp:keywords>
    <cp:lastModifiedBy>spdocx</cp:lastModifiedBy>
    <cp:revision>1</cp:revision>
</cp:coreProperties>
)";

static const char * s_rels_rels = R"(<?xml encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
)";

static const char * s_word_document = R"(<?xml encoding="UTF-8" standalone="yes"?>
<w:document 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" 
    mc:Ignorable="w14 w15">
    <w:body>
        <w:p w:rsidR="00C35A5F" w:rsidRDefault="00C35A5F" w:rsidP="00C35A5F"/>
    </w:body>
</w:document>
)";

static const char * s_word_styles = R"(<?xml encoding="UTF-8" standalone="yes"?>
<w:styles 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" 
    mc:Ignorable="w14 w15">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                <w:kern w:val="2"/>
                <w:sz w:val="21"/>
                <w:szCs w:val="22"/>
                <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault/>
    </w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="a">
        <w:name w:val="Normal"/>
        <w:qFormat/>
        <w:rsid w:val="00763DC6"/>
        <w:pPr>
            <w:widowControl w:val="0"/>
            <w:jc w:val="both"/>
        </w:pPr>
        <w:rPr>
            <w:szCs w:val="21"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="a"/>
        <w:next w:val="20"/>
        <w:link w:val="10"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:rsid w:val="00A2258A"/>
        <w:pPr>
            <w:keepNext/>
            <w:keepLines/>
            <w:numPr>
                <w:numId w:val="22"/>
            </w:numPr>
            <w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:eastAsia="ºÚÌå" w:hAnsi="Times New Roman"/>
            <w:b/>
            <w:bCs/>
            <w:kern w:val="44"/>
            <w:sz w:val="44"/>
            <w:szCs w:val="44"/>
        </w:rPr>
    </w:style>
</w:styles>
)";
// TODO : more style,

// TODO : numbering.xml

static const char * s_word_rels = R"(<?xml encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
</Relationships>
)";

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

	m_files["[Content_Types].xml"].assign( s_content_type, s_content_type + strlen( s_content_type ) );
	m_files["docProps/app.xml"].assign( s_docprops_app, s_docprops_app + strlen( s_docprops_app ) );
	m_files["docProps/core.xml"].assign( s_docprops_core, s_docprops_core + strlen( s_docprops_core ) );
	m_files["_rels/.rels"].assign( s_rels_rels, s_rels_rels + strlen( s_rels_rels ) );
	m_files["word/document.xml"].assign( s_word_document, s_word_document + strlen( s_word_document ) );
	m_files["word/styles.xml"].assign( s_word_styles, s_word_styles + strlen( s_word_styles ) );
	m_files["word/_rels/document.xml.rels"].assign( s_word_styles, s_word_styles + strlen( s_word_styles ) );

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
	
	// TODO : style and rela

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
			if( ppr ) {
				pugi::xml_node numid = ppr.child( "w:numPr" );
				if( numid )
					numid = numid.child( "w:numId" );
				if( numid ) {
					style.m_pnumid = numid.attribute( "w:val" ).value();
				}
				else {
					style.m_pnumid = "";
				}
			}
			else {
				style.m_pnumid = "";
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
