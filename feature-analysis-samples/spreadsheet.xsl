<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  >
	<xsl:template match="/" >
		<html>
			<head>
				Transformed Spreadsheet
			</head>

			<script src="spreadSheetFormulae.js"></script>

			<body>
			<style type="text/css">

.excelDefaults {
	background-color: white;
	color: black;
	text-decoration: none;
	direction: ltr;
	text-transform: none;
	text-indent: 0;
	letter-spacing: 0;
	word-spacing: 0;
	white-space: normal;
	unicode-bidi: normal;
	vertical-align: 0;
	background-image: none;
	text-shadow: none;
	list-style-image: none;
	list-style-type: none;
	padding: 0;
	margin: 0;
	border-collapse: collapse;
	white-space: pre;
	vertical-align: bottom;
	font-style: normal;
	font-family: sans-serif;
	font-variant: normal;
	font-weight: normal;
	font-size: 10pt;
	text-align: right;
}

.excelDefaults td {
	padding: 1px 5px;
	border: 1px solid silver;
}


.excelDefaults .colHeader {
	background-color: silver;
	font-weight: bold;
	border: 1px solid black;
	text-align: center;
	padding: 1px 5px;
}
.excelDefaults .formula {
	border: 1px solid orange;
}
.excelDefaults .rowHeader {
	background-color: silver;
	font-weight: bold;
	border: 1px solid black;
	text-align: right;
	padding: 1px 5px;
}
</style>
	<xsl:for-each select="spreadsheets/Table">
		
				<table  class="excelDefaults">
					<tr class="colHeader">
						<xsl:for-each select="ColumnHeaders/ColumnHeader">
							<th class="colHeader">
								<xsl:value-of select=".">
								</xsl:value-of>
							</th>
						</xsl:for-each>
					</tr>
					<xsl:for-each select="TableRow">
						<tr>
										<td class="rowHeader">
											<xsl:value-of select="RowHeader"/>
										</td>

							<xsl:for-each select="TableCells/TableCell">								
										<td>
											<xsl:element name="input">
												<xsl:attribute name="type">
                          text
                        </xsl:attribute>
												<xsl:attribute name="value">
                          <xsl:value-of select="@value">
                          </xsl:value-of>
                        </xsl:attribute>
												<xsl:if test="@formula">
													<xsl:attribute name="onchange">
                            javascript:ApplyFormula(this, '<xsl:value-of
														select="@formula"></xsl:value-of>');
                          </xsl:attribute>
												</xsl:if>
												
						<xsl:if test="@cellFormula">
													<xsl:attribute name="class">
	                            formula
                          </xsl:attribute>
                          	<xsl:attribute name="onmouseover">
                            javascript:document.getElementById('formulaShow').innerHTML= '<xsl:value-of
														select="@cellFormula"></xsl:value-of>'
                            
                          </xsl:attribute>
                          	<xsl:attribute name="onmouseout">
                          	 javascript:document.getElementById('formulaShow').innerHTML= ''
                          	</xsl:attribute>
												</xsl:if>
												<xsl:attribute name="id">
                          <xsl:value-of select="@cellID">
                          </xsl:value-of>
                        </xsl:attribute>
												<xsl:if test="@readOnly">
													<xsl:attribute name="readonly">
                            readonly
                          </xsl:attribute>
												</xsl:if>
											</xsl:element>
										</td>

							</xsl:for-each>
						</tr>
					</xsl:for-each>
				</table>
				<br/>
				<br/>
			</xsl:for-each>
			Formula: <div id="formulaShow"></div>
			
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>
