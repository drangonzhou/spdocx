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
#include <string>

#include <time.h>

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

using namespace spd;

static void dump_element( const Document * doc, Element ele, int level )
{
	static const char pre[16] = "               ";
	const char * p = ( level >= 15 ) ? pre : pre + 15 - level;
	for( ; ele.GetType() != ElementTypeE::INVALID; ele = ele.GetNext() ) {
		printf( "%s[%s] (%d)\n", p, ele.GetTag(), (int)ele.GetType() );
		switch( ele.GetType() )
		{
		case ElementTypeE::PARAGRAPH :
		{
			Paragraph & par = static_cast< Paragraph & >( ele );
			printf( "<style> [%s]\n", par.GetStyleName( doc ) );
			printf( "<text> [%s]\n", par.GetText().c_str() );
		}
			break;
		case ElementTypeE::HYPERLINK :
		{
			Hyperlink & hlk = static_cast< Hyperlink & >( ele );
			printf( "<anchor> [%s]\n", hlk.GetAnchor() );
			printf( "<linkType> [%s]\n", hlk.GetLinkType( doc ) );
			printf( "<target> mode [%s] : [%s]\n", hlk.GetTargetMode( doc ), hlk.GetTarget( doc ) );
			printf( "<text> [%s]\n", hlk.GetText().c_str() );
		}
			break;
		case ElementTypeE::RUN :
		{
			Run & run = static_cast< Run & >( ele );
			printf( "<color> [%s], <hightline> [%s]\n", run.GetColor(), run.GetHighline() );
			printf( "<bold> [%s], <italic> [%s], <underline> [%s], <strike> [%s], <dstrike> [%s]\n", 
				run.GetBold() ? "true" : "false",
				run.GetItalic() ? "true" : "false",
				run.GetUnderline(), 
				run.GetStrike() ? "true" : "false",
				run.GetDoubleStrike() ? "true" : "false" );
			printf( "<text> [%s]\n", run.GetText().c_str() );
		}
			break;
		case ElementTypeE::TABLE :
		{
			Table & tbl = static_cast< Table & >( ele );
			printf( "<table> Row %d, Col %d, width [", tbl.GetRowNum(), tbl.GetColNum() );
			for( int i = 0; i < tbl.GetColNum(); ++i ) {
				if( i != 0 )
					printf( "," );
				printf( " %d", tbl.GetColWidth( i ) );
			}
			printf( " ]\n" );
		}
			break;
		case ElementTypeE::TABLE_ROW :
			printf( "<table row>\n" );
			break;
		case ElementTypeE::TABLE_CELL :
		{
			TCell & tc = static_cast< TCell & >( ele );
			printf( "<table cell> span %d, vmerge %d\n", tc.GetSpanNum(), (int)tc.GetVMergeType() );
			printf( "<text> [%s]", tc.GetText().c_str() );
		}
			break;
		default:
			break;
		}
		dump_element( doc, ele.GetFirstChild(), level + 1 );
	}
	return;
}

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

	Element ele = doc.GetFirstElement();

	Element ele2 = ele.GetParent();
	printf( "parent is %d\n", (int)ele2.GetType() );

	dump_element( &doc, ele, 0 );

	ele = Element(); // nullptr
	doc.Close();

	ret = doc.Open( fname );
	if( ret < 0 ) {
		printf( "doc.Open() ret %d\n", ret );
	}

	ele = doc.GetFirstElement();
	ele = ele.GetFirstChild();
	if( ele.GetType() == ElementTypeE::RUN ) {
		Run & run = static_cast< Run &>( ele );
		// TODO : test
		int tt = (int)time( NULL );
		std::string str = std::to_string( tt );
		printf( "before modify : [%s]\n", run.GetText().c_str() );
		run.SetText( str.c_str() );
		printf( "after modify : [%s]\n", run.GetText().c_str() );
	}

	ele = Element(); // nullptr;
	doc.Save( "t2.docx" );

	SPD_PR_INFO( "hello %d", 5 );

	return 0;
}
