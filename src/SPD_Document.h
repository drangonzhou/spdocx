// SPD_Document.h : spdocx document
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

#ifndef INCLUDED_SPD_DOCUMENT_H
#define INCLUDED_SPD_DOCUMENT_H

#include "SPD_Common.h"

#include "zip.h"
#include "pugixml.hpp"

BEGIN_NS_SPD
////////////////////////////////

class SPD_API Document
{
public:
	Document();
	~Document();

	int Open( const char * fname );
	int Save( const char * fname = NULL );
	int Close();

private:
	zip_t * m_zip;
	pugi::xml_document m_doc;
};

////////////////////////////////
END_NS_SPD

#endif // INCLUDED_SPD_DOCUMENT_H
