
# 1 docx 文档概述

这里通过简要的例子描述一下 docx 文档的结构，省略大部分信息，完整内容请参考 OpenXML 规范。

docx 文档本身是一个 zip 压缩包，里面通过多个文件描述文档的内容、样式、内嵌图片等。

# 2 文档内容 （ word/document.xml ）

## 2.1 整体结构（ w:document/w:body ）

```
<?xml version="1.0"?>
<w:document ...>
	<w:body>
		<w:p> ... </w:p>
		<w:tbl> ... </w:tbl>
		...
		<w:sectPr> ... <w:sectPr>
	</w:body>
</w:document>
```

## 2.2 文本段落（ w:p ）

```
<w:p>
	<w:pPr>
		<w:pStyle w:val="1" />    <!-- 样式，引用 word/styles.xml -->
		<w:jc w:val="center" />   <!-- 对齐方式，取值："left", "center", "right", "both" -->
		<w:numPr>
			<w:numId w:val="2" />  <!-- 列表信息，引用 word/numbering.xml -->
			<w:ilvl w:val="1" />   <!-- 列表层级，引用 word/numbering.xml -->
		</w:numPr>
	</w:pPr>
	<w:r> ... </w:r>
	<w:hyperlink> ... </w:hyperlink>
	<w:bookmarkStart> ... </w:bookmarkStart>
	<w:bookmarkEnd> ... </w:bookmarkEnd>
	...
</w:p>
```
关联关系：
* w:pStyle 属性 w:val="1" 关联 word/styles.xml 的 w:style 属性 w:styleId="1" 
* w:numId 属性 w:val="2" 关联 word/numbering.xml 的 w:num 属性 w:numId="1"
* w:ilvl 属性 w:val="1" 关联 word/numbering.xml 的 w:abstractNum 下的 w:lvl 属性 w:ilvl="1"

## 2.3 文本图像等内容块（ w:r ）

### 2.3.1 文本块（ w:t ）

```
<w:r>
	<w:rPr>
		<w:b />   <!-- 粗体 -->
		<w:i />   <!-- 斜体 -->
		<w:u w:val="single" />   <!-- 下划线，取值："single", "double", "dotted" -->
		<w:strike />    <!-- 删除线 -->
		<w:dstrike />   <!-- 双删除线 -->
		<w:color w:val="FF0000"/>    <!-- 颜色 -->
		<w:highlight w:val="yellow"/>    <!-- 高亮，背景色 -->
	</w:rPr>
	<w:t>这是文本</w:t>
</w:r>
```

### 2.3.2 图像块（ w:drawing ）

```
<w:r>
	<w:drawing>
		<wp:anchor>
			<a:graphic>
				<a:graphicData>
					<pic:pic>
						<pic:blipFill>
							<a:blip r:embed="rId4">  <!-- 图片，引用 word/_rels/document.xml.rels -->
							...
							</a:blip>
						</pic:blipFill>
					</pic:pic>
				</a:graphicData>
			</a:graphic>
		</wp:anchor>
		<wp:inline> ... </wp:inline>
	</w:drawing>
</w:r>
```

通过 r:embed 关联的 document.xml.rels 里面的 Relationship 的 Id ：
```
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/example.png"/>
```

### 2.3.3 嵌入对象（ w:object ）

嵌入 visio 对象

```
<w:r>
	<w:object>
		<v:shapetype> ... </v:shapetype>
		<v:shape>
			<v:imagedata r:id="rId5" />   <!-- 图片，emf图片，引用 word/_rels/document.xml.rels -->
		</v:shape>
		<o:OLEObject Type="Embed" ProgID="Visio.Drawing.15" r:id="rId6"/>   <!-- 嵌入对象，visio文件，引用 word/_rels/document.xml.rels -->
	</w:object>
</w:r>
```

visio 文件及其图片通过 r:id 关联的 document.xml.rels 里面的 Id ：
```
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Visio___.vsdx"/>
```

## 2.4 超链接（ w:hyperlink ）

指向外部超链接

```
<w:hyperlink r:id="rId2">  <!-- 外部超链接，引用word/_rels/document.xml.rels -->
	<w:r> ... </w:r>
	...
</w:hyperlink>
```

指向内部书签

```
<w:hyperlink w:anchor="myBookmark" >  <!-- 内部书签，引用w:bookmarkStart -->
	<w:r> ... </w:r>
	...
</w:hyperlink>
```

## 2.5 表格（ w:tbl ）

### 2.5.1 表格整体

```
<w:tbl>
	<w:tblGrid>
		<w:gridCol w:w="2088" />   <!-- 列宽度，页总宽大约8000多 -->
		...
	</w:tblGrid>
	<w:tr> ... </w:tr>
	...
</w:tbl>
```

### 2.5.2 表格的行（ w:tr ）

```
<w:tr>
	<w:tc> ... </w:tc>
	...
</w:tr>
```

### 2.5.3 表格的单元（ w:tc ）

```
<w:tc>
	<w:tcPr>
		<w:gridSpan w:val="2" />        <!-- 横向单元格合并个数 -->
		<w:vMerge w:val="restart" />    <!-- 纵向单元格合并，取值 "restart" "continue" -->
	</w:tcPr>
	<w:p> ... </w:p>
	<w:tbl> ... </w:tbl>
	...
</w:tc>
```

### 2.6 书签（ w:bookmarkStart，w:bookmarkEnd ）

书签元素在 w:p 内，但开始结束可以跨多个 w:p ，引用时通过 w:hyperlink 引用

```
<w:p>
	<w:bookmarkStart w:id="0" w:name="myBookmark" />
	...
</w:p>
<w:p>
	...
	<w:bookmarkEnd  w:id="0" />
</w:p>
```

# 3 文本样式 （ word/style.xml ）

# 3.1 整体结构

```
<?xml version="1.0"?>
<w:styles>
	<w:style> ... </w:style>
	...
</w:styles>
```

# 3.2 样式（ w:style ）

```
<w:style w:type="paragraph" w:styleId="1" >
	<w:name w:val="heading 1"/>     <!-- 样式名称，标题和正文等  -->
	<w:pPr>
		<w:numPr>
			<w:numId w:val="22"/>   <!-- 列表信息，引用 word/numbering.xml  -->
			<w:ilvl w:val="1" />   <!-- 列表层级，引用 word/numbering.xml -->
		</w:numPr>
	</w:pPr>
</w:style>
```

# 4 引用关系 （ word/_rels/document.xml.rels ）

# 4.1 整体结构

```
<?xml version="1.0"?>
<Relationships>
	<Relationship> ... </Relationship>
	...
</Relationships>
```

# 4.2 关系描述（ Relationship ）

```
<Relationship Id="rId1" 
	Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" 
	Target="http://www.example.com" TargetMode="External" />
```

常见类型：
* 外部超链接： http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink
* 图片： http://schemas.openxmlformats.org/officeDocument/2006/relationships/image

# 5 列表信息 （ word/numbering.xml ）

## 5.1 整体结构

```
<?xml version="1.0"?>
<w:numbering>
	<w:abstractNum> ... </w:abstractNum>
	<w:num> ... </w:num>
	...
</w:numbering>
```

## 5.2 列表模版 （ w:abstractNum ）

```
<w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0" >
	<w:multiLevelType w:val="multilevel"/>
	<w:numStyleLink w:val="XXX"/>     <!-- 表示引用另一个模版  -->
	<w:styleLink w:val="XXX"/>        <!-- 表示可以被引用  -->
	<w:lvl w:ilvl="1" >
		<w:start w:val="1"/>
		<w:numFmt w:val="chineseCountingThousand"/>
		...
	</w:lvl>
	...
</w:abstractNum>
```


## 5.3 列表实例 （ w:num ）

```
<w:num w:numId="1">
	<w:abstractNumId w:val="1" />   <!-- 引用模版  -->
</w:num>
```

