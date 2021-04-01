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

	int err = 0;
	m_zip = zip_open( fname, 0, &err );
	if( err < 0 )
		return 0;
	return 0;
}

int Document::Save( const char * fname )
{
	return -1;
}

int Document::Close()
{
	if( m_zip != nullptr ) {
		zip_discard( m_zip ), m_zip = nullptr;
	}
	m_doc.reset();
	return 0;
}

////////////////////////////////
END_NS_SPD
