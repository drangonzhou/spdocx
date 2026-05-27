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
#include <fstream>
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

static void dump_element( const Document & doc, Element ele, int level )
{
	static const char pre[16] = "               ";
	const char * p = ( level >= 15 ) ? pre : pre + 15 - level;
	int cell = 0;
	int col = 0;
	int row = 0;
	for( ; ele.GetType() != ElementTypeE::INVALID; ele = ele.GetNext() ) {
		printf( "%s[%s] (%d)\n", p, ele.GetTag(), (int)ele.GetType() );
		switch( ele.GetType() )
		{
		case ElementTypeE::PARAGRAPH :
		{
			Paragraph par = ele;
			printf( "<style> [%s]\n", par.GetStyleName( doc ) );
			printf( "<text> [%s]\n", par.GetText().c_str() );
		}
			break;
		case ElementTypeE::HYPERLINK :
		{
			Hyperlink hlk = ele;
			printf( "<Anchor> [%s]\n", hlk.GetAnchor() ? hlk.GetAnchor() : "<nullptr>" );
			printf( "<RelaId> [%s]\n", hlk.GetRelaId() ? hlk.GetRelaId() : "<nullptr>" );
			if( hlk.GetRelaId() != nullptr ) {
				printf( "<RelaType> [%s]\n", hlk.GetRelaType( doc ).c_str() );
				printf( "<TargetMode> [%s]\n", hlk.GetRelaTargetMode( doc ).c_str() );
				printf( "<Target> [%s]\n", hlk.GetRelaTarget( doc ).c_str() );
			}
			printf( "<text> [%s]\n", hlk.GetText().c_str() );
		}
			break;
		case ElementTypeE::RUN :
		{
			Run run = ele;
			if( run.IsText() ) {
				printf( "<color> [%s], <hightline> [%s]\n", run.GetColor(), run.GetHighline() );
				printf( "<bold> [%s], <italic> [%s], <underline> [%s], <strike> [%s], <dstrike> [%s]\n",
					run.GetBold() ? "true" : "false",
					run.GetItalic() ? "true" : "false",
					run.GetUnderline(),
					run.GetStrike() ? "true" : "false",
					run.GetDoubleStrike() ? "true" : "false" );
				printf( "<text> [%s]\n", run.GetText().c_str() );
			}
			else if( run.IsPic() ) {
				printf( "<PicId> [%s]\n", run.GetPicRelaId() ? run.GetPicRelaId() : "<nullptr>" );
			}
			else if( run.IsObject() ) {
				printf( "<ObjectId> [%s]\n", run.GetObjectRelaId() ? run.GetObjectRelaId() : "<nullptr>" );
				printf( "<ProgId> [%s]\n", run.GetObjectProgId() ? run.GetObjectProgId() : "<nullptr>" );
				printf( "<ImgId> [%s]\n", run.GetObjectImgRelaId() ? run.GetObjectImgRelaId() : "<nullptr>" );
			}
		}
			break;
		case ElementTypeE::TABLE :
		{
			Table tbl = ele;
			row = 0;
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
			printf( "<table row %d>\n", row++ );
			cell = 0;
			col = 0;
			break;
		case ElementTypeE::TABLE_CELL :
		{
			TCell tc = ele;
			printf( "<table cell %d (col %d)> span %d, vmerge %d\n", cell, col, tc.GetSpanNum(), (int)tc.GetVMergeType() );
			++cell;
			col += tc.GetSpanNum();
			printf( "<text> [%s]", tc.GetText().c_str() );
		}
			break;
		case ElementTypeE::BOOKMARK_END :
			printf( "<bookmark end>\n" );
			break;
		default:
			break;
		}
		dump_element( doc, ele.GetFirstChild(), level + 1 );
	}
	return;
}


static int dumpinfo( const char * fname )
{
	int ret;
	Document doc;
	ret = doc.Open( fname );
	if( ret < 0 ) {
		printf( "doc.Open() ret %d\n", ret );
		return -1;
	}

	SPDDebug::DumpDocument( &doc );

	Element ele = doc.GetFirstElement();

	Element ele2 = ele.GetParent();
	printf( "parent is %d\n", (int)ele2.GetType() );

	dump_element( doc, ele, 0 );

	ele = Element(); // nullptr
	doc.Close();

	ret = doc.Open( fname );
	if( ret < 0 ) {
		printf( "doc.Open() ret %d\n", ret );
	}

	ele = doc.GetFirstElement();
	ele = ele.GetFirstChild();
	if( ele.GetType() == ElementTypeE::RUN ) {
		Run run = ele;
		// TODO : test
		int tt = (int)time( NULL );
		std::string str = std::to_string( tt );
		printf( "before modify : [%s]\n", run.GetText().c_str() );
		run.SetText( str.c_str() );
		printf( "after modify : [%s]\n", run.GetText().c_str() );
	}

	ele = Element(); // nullptr;
	doc.Save( "t2.docx" );

	return 0;
}

static int dumpdirjson( const char * fname )
{
	int ret;
	Document doc;
	ret = doc.Open( fname );
	if( ret < 0 ) {
		printf( "doc.Open() ret %d\n", ret );
	}
	for( auto & ele : doc )
	{
		if( ele.GetType() != ElementTypeE::PARAGRAPH )
			continue;

	}
	doc.Close();
	return 0;
}

static int newdoc( const char * fname )
{
	int ret;
	Document doc;
	doc.New();
	// title
	Element ele = doc.GetFirstElement();
	ele.DelAllChild();
	Paragraph para = ele;
	Run r = para.AddChildRun();
	r.SetText( "this is title" );
	Paragraph para2 = para.AddSiblingParagraph();
	para2.SetStyleId( doc.GetStyleId( "heading 1" ) );
	Run r2 = para2.AddChildRun();
	r2.SetText( "chapter one" );
	ret = doc.Save( fname );
	if( ret < 0 ) {
		printf( "doc.Save() ret %d\n", ret );
	}
	for( auto & ele : doc )
	{
		if( ele.GetType() != ElementTypeE::PARAGRAPH )
			continue;

	}
	doc.Close();
	return 0;
}

static int t_table()
{
	int ret;
	const char * fname = "_test_table.docx";

	// 1. create doc with table
	{
		Document doc;
		doc.New();

		// clear auto-created paragraph
		doc.GetFirstElement().DelAllChild();

		// add a 5x5 table
		Table tbl = doc.AddChildTable();
		if( !tbl.IsValid() ) {
			printf( "t_table: AddChildTable FAILED!\n" );
			return -1;
		}
		ret = tbl.Reset( 5, 5 );
		if( ret < 0 ) {
			printf( "t_table: Reset FAILED ret=%d\n", ret );
			return -1;
		}

		// fill each cell with text label [row, col]
		int ri = 0;
		for( TRow row = tbl.GetFirstChild(); row.IsValid(); row = row.GetNext(), ++ri )
		{
			for( TCell cell = row.GetFirstChild(); cell.IsValid(); cell = cell.GetNext() )
			{
				if( cell.GetType() == ElementTypeE::TABLE_CELL )
				{
					Paragraph p = cell.AddChildParagraph();
					Run r = p.AddChildRun();
					char buf[64];
					snprintf( buf, sizeof(buf), "R%dC%d", ri, cell.GetCol() );
					r.SetText( buf );
				}
			}
		}

		// horizontal merge (span): merge row 0, col 1 and col 2 (col 1 span 2)
		{
			TRow row0 = tbl.GetFirstChild();
			if( !row0.IsValid() ) {
				printf( "t_table: GetFirstChild row0 FAILED!\n" );
				return -1;
			}
			TCell cell = row0.GetCell( 1 );
			if( !cell.IsValid() ) {
				printf( "t_table: GetCell(1) FAILED!\n" );
				return -1;
			}
			ret = cell.SetSpanNum( 2 );
			if( ret < 0 ) {
				printf( "t_table: SetSpanNum FAILED ret=%d\n", ret );
				return -1;
			}
		}

		// vertical merge (vmerge): merge row 1 col 0, row 2 col 0, row 3 col 0 (cell row1 col0 becomes VMerge START spanning 3)
		{
			TRow row0 = tbl.GetFirstChild();
			TRow row1 = row0.GetNext();
			if( !row1.IsValid() ) {
				printf( "t_table: GetNext row1 FAILED!\n" );
				return -1;
			}
			TCell cell = row1.GetCell( 0 );
			if( !cell.IsValid() ) {
				printf( "t_table: GetCell(0) for vmerge FAILED!\n" );
				return -1;
			}
			ret = cell.SetVMergeNum( 3 );
			if( ret < 0 ) {
				printf( "t_table: SetVMergeNum FAILED ret=%d\n", ret );
				return -1;
			}
		}

		ret = doc.Save( fname );
		doc.Close();
		printf( "t_table: saved [%s] ret=%d\n", fname, ret );
	}

	// 2. read back and verify
	{
		Document doc;
		ret = doc.Open( fname );
		if( ret < 0 )
		{
			printf( "t_table: open [%s] FAILED ret=%d\n", fname, ret );
			return -1;
		}

		Table tbl = Element(); // default-construct Element then convert to Table
		for( Element ele = doc.GetFirstElement(); ele.IsValid(); ele = ele.GetNext() )
		{
			if( ele.GetType() == ElementTypeE::TABLE )
			{
				tbl = Table( ele );
				break;
			}
		}
		if( !tbl.IsValid() )
		{
			printf( "t_table: first element is not TABLE, type=%d\n", (int)tbl.GetType() );
			return -1;
		}

		printf( "t_table: table %dx%d\n", tbl.GetRowNum(), tbl.GetColNum() );
		if( tbl.GetRowNum() < 4 || tbl.GetColNum() < 5 ) {
			printf( "t_table: table too small (%dx%d), expected at least 4x5\n",
				tbl.GetRowNum(), tbl.GetColNum() );
			return -1;
		}

		int err = 0;

		// verify horizontal merge: row 0, col 1 should have span 2
		{
			TRow row0 = tbl.GetFirstChild();
			TCell cell = row0.GetCell( 1 );
			if( !cell.IsValid() ) {
				printf( "t_table: verify GetCell(1) FAILED!\n" );
				return -1;
			}
			int span = cell.GetSpanNum();
			printf( "t_table: row0 col1 span=%d (expected 2)\n", span );
			if( span != 2 ) err++;
		}

		// verify vertical merge: row 1 col 0 should be VMerge START with num 3
		// row 2 col 0 should be VMerge CONT
		// row 3 col 0 should be VMerge CONT
		{
			TRow row0 = tbl.GetFirstChild();
			TRow row1 = row0.GetNext();
			if( !row1.IsValid() ) {
				printf( "t_table: verify row1 FAILED!\n" );
				return -1;
			}
			TRow row2 = row1.GetNext();
			if( !row2.IsValid() ) {
				printf( "t_table: verify row2 FAILED!\n" );
				return -1;
			}
			TRow row3 = row2.GetNext();
			if( !row3.IsValid() ) {
				printf( "t_table: verify row3 FAILED!\n" );
				return -1;
			}

			TCell c1 = row1.GetCell( 0 );
			TCell c2 = row2.GetCell( 0 );
			TCell c3 = row3.GetCell( 0 );
			if( !c1.IsValid() || !c2.IsValid() || !c3.IsValid() ) {
				printf( "t_table: verify GetCell for vmerge FAILED!\n" );
				return -1;
			}

			int vnum1 = c1.GetVMergeNum();
			VMergeTypeE vt1 = c1.GetVMergeType();
			VMergeTypeE vt2 = c2.GetVMergeType();
			VMergeTypeE vt3 = c3.GetVMergeType();

			printf( "t_table: row1 col0 vmergeNum=%d type=%d (expected 3, START=%d)\n", vnum1, (int)vt1, (int)VMergeTypeE::START );
			printf( "t_table: row2 col0 vmergeType=%d (expected CONT=%d)\n", (int)vt2, (int)VMergeTypeE::CONT );
			printf( "t_table: row3 col0 vmergeType=%d (expected CONT=%d)\n", (int)vt3, (int)VMergeTypeE::CONT );

			if( vnum1 != 3 || vt1 != VMergeTypeE::START || vt2 != VMergeTypeE::CONT || vt3 != VMergeTypeE::CONT )
				err++;
		}

		doc.Close();

		if( err > 0 )
		{
			printf( "t_table: VERIFY FAILED with %d error(s)!\n", err );
		}
		else
		{
			printf( "t_table: VERIFY OK!\n" );
		}
	}

	// 3. delete test file
	remove( fname );
	printf( "t_table: deleted [%s]\n", fname );

	return 0;
}

static int conv( const char * fname )
{
	// read file to char string
	std::ifstream ifs( fname, std::ios::binary );
	if( !ifs.is_open() ) {
		printf( "can't open file %s\n", fname );
		return -1;
	}
	ifs.seekg( 0, std::ios::end );
	long len = ifs.tellg();
	ifs.seekg( 0, std::ios::beg );
	char * buf = new char[len + 1];
	ifs.read( buf, len );
	buf[len] = '\0';
	ifs.close();

	printf( "const char * const str = \n\"" );
	long i;
	for( i = 0; i < len; ++i )
	{
		uint8_t ch = (uint8_t)buf[i];
		if( ch == '"' ) {
			printf( "\\\"" );
		}
		else if( ch == '\r' ) {
			printf( "\\r" );
		}
		else if( ch == '\n' ) {
			printf( "\\n" );
		}
		else if( ch >= 32 && ch <= 126 ) {
			printf( "%c", ch );
		}
		else {
			printf( "\\x%02x", ch );
		}
		if( i % 120 == 119 )
		{
			printf( "\"\n\"" );
		}
	}
	printf( "\";\n" );

	delete [] buf, buf = nullptr;
	return 0;
}

static int usage()
{
	printf( "%s", R"(Usage: spdocxutil <cmd> <params>
  usage | help        : show this usage
  dumpinfo            : dump docx info
  dumpdirjson         : dump docx dir in json
  newdoc              : create new docx 
  conv                : conv file to char string
  t_table             : test table create/merge/verify
)" );
	return 0;
}

int main( int argc, char * argv[] )
{
	if( argc < 2 ) {
		usage();
		return -1;
	}

	spd::SPD_SetLogLevel( SPD_LOG_LEVEL_DEBUG );
	SPD_PR_INFO( "hello %d", 5 );

	if( strcmp( argv[1], "dumpinfo" ) == 0 && argc >= 3 ) {
		dumpinfo( argv[2] );
		//char utf8_string[] = "\xE4\xB8\xAD\xE6\x96\x87.docx"; // 中文.docx
		//dumpinfo( utf8_string );
	}
	else if( strcmp( argv[1], "dumpdirjson" ) == 0 && argc >= 3 ) {
		dumpdirjson( argv[2] );
	}
	else if( strcmp( argv[1], "newdoc" ) == 0 && argc >= 3 ) {
		newdoc( argv[2] );
	}
	else if( strcmp( argv[1], "conv" ) == 0 && argc >= 3 ) {
		conv( argv[2] );
	}
	else if( strcmp( argv[1], "t_table" ) == 0 ) {
		t_table();
	}
	else {
		printf( "[ERR] unknown cmd or bad params\n" );
		usage();
	}

	return 0;
}
