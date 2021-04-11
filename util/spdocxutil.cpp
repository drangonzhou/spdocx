// spdocxutil.cpp : spdocx util
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

#include <iostream>

using namespace spd;

BEGIN_NS_SPD
////////////////////////////////

class SPDDebug
{
public:
    static void DumpDocument( Document * doc );
};

void SPDDebug::DumpDocument( Document * doc )
{
    printf( "[doc] %p\n", doc );
    printf( "  [zip] %p\n", doc->m_zip );
    for( auto it = doc->m_style.begin(); it != doc->m_style.end(); ++it ) {
        printf( "  [style] id [%s], name [%s], type [%s]\n", 
            it->second.m_id.c_str(), it->second.m_name.c_str(), it->second.m_type.c_str() );
    }
    for( auto it = doc->m_rela.begin(); it != doc->m_rela.end(); ++it ) {
        printf( "  [rela] id [%s], type [%s], target [%s], targetMode [%s]\n", 
            it->second.m_id.c_str(), it->second.m_type.c_str(), it->second.m_target.c_str(), it->second.m_targetMode.c_str() );
    }
    return;
}

////////////////////////////////
END_NS_SPD

int main( int argc, char * argv[] )
{
    if( argc < 2 ) {
        printf( "Usage: %s <docx file>\n", argv[0] );
        return -1;
    }

    spd::SPD_SetLogFuncLevel( nullptr, SPD_LOG_LEVEL_DEBUG );

    const char * fname = argv[1];
    int ret;
    Document doc;
    ret = doc.Open( fname );
    if( ret < 0 ) {
        printf( "doc.Open() ret %d\n", ret );
    }

    SPDDebug::DumpDocument( &doc );

    doc.Close();

    SPD_PR_INFO( "hello %d", 5 );

    return 0;
}
