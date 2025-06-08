// SPD_Document.h : spdocx document
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

#ifndef INCLUDED_SPD_DOCUMENT_H
#define INCLUDED_SPD_DOCUMENT_H

#include "SPD_Common.h"
#include "SPD_Element.h"

#include <pugixml.hpp>
#include <string>
#include <map>
#include <vector>

BEGIN_NS_SPD
////////////////////////////////

class SPD_API StyleLite
{
public:
	std::string m_id;
	std::string m_type;  // paragraph, table, character ...
	std::string m_name;  // "headding 1", "heading 2", ... , "normal",
	std::string m_numId; // empty means no numbering
	int m_numLevel = -1;  // level start from 0, -1 means no numbering, 
};

class SPD_API Relationship
{
public:
	std::string m_id;
	std::string m_type;        // "hyperlink", "image", "media", ...
	std::string m_target;      // name or url
	std::string m_targetMode;  // External
};

class SPD_API Document
{
public:
	Document();
	~Document();

	int New();
	int Open( const std::string & fname );
	int Save( const std::string & fname = std::string{} );
	int Close();
	bool IsValid() const { return ! m_files.empty(); }
	bool IsModified() const { return m_isModified; }
	std::string GetFileName() const { return m_fname; }

	// Paragraph or Table, or other Element ( Section )
	// If no Element, auto create an empty Paragraph
	Element GetFirstElement() const;
	ElementIterator begin() const { return ElementIterator( GetFirstElement() ); }
	ElementIterator end() const { return ElementIterator(); }
	ElementRange Children() const { return ElementRange( begin(), end() ); }

	// default is add to back, 
	Paragraph AddChildParagraph( bool add_back = true );
	Table AddChildTable( bool add_back = true );
	int DelChild( Element & child ); // child must be direct child
	int DelAllChild();

	const char * GetStyleName( const char * id ) const;
	const char * GetStyleId( const char * name ) const;
	const StyleLite * GetStyle( const char* id ) const;
	std::vector< const StyleLite *> GetAllStyle() const;
	int AddStyle( const StyleLite& style );  // if style exist, update it
	const Relationship * GetRelationship( const char * id ) const;
	std::vector< const Relationship * > GetAllRelationship() const;
	int AddRelationship( const Relationship & rela );  // if relationship exist, update it

	int GetEmbedData( const std::string & id, std::vector<char> & data ) const;  // get embedding data in "word/" directly, ex : media/image1.png
	// TODO : set embed data

protected:
	static int read_zip( const std::string & fname, std::map< std::string, std::vector<char> > & files );
	static int write_zip( const std::string & fname, const std::map< std::string, std::vector<char> > & files );

	int parse_files();
	int read_xml( const std::string & fname, pugi::xml_document * doc ) const;
	int write_xml( const std::string & fname, const pugi::xml_document & doc );
	int load_style();
	int load_rela();
	int write_style();
	int write_rela();

private:
	friend class SPDDebug;
	std::string m_fname;
	std::map< std::string, std::vector<char> > m_files;  // zip files, every file has an extra '\0'
	bool m_isModified = false;
	
	pugi::xml_document m_doc;
	std::map< std::string, StyleLite > m_style;
	std::map< std::string, Relationship > m_rela;
};

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_DOCUMENT_H
