<?xml version="1.0" encoding="UTF-8"?>
<!-- ***************************************************************************
 $Revision: 1.2 $
 $Date: 2010/07/23 16:56:53 $
 Author: doudou
 Description: Template for AU3 constant exporter.
 *************************************************************************** -->
<xsl:stylesheet version="2.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:fn="http://www.w3.org/2005/xpath-functions"
    xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
    xmlns:p="http://autoitscript/tldoc/xslt">
    
    <xsl:output
        omit-xml-declaration="yes"
        encoding="UTF-8"
        indent="no"
        method="text"
        media-type="text/plain"
        exclude-result-prefixes="xsl xs fn xdt p"/>
    
    <xsl:param name="scope">Global</xsl:param>
    <xsl:param name="indent" xml:space="preserve">    </xsl:param>
    
    <!-- entry point -->
    <xsl:template match="/">
        <xsl:apply-templates select="TypeLib/Types/Enum" mode="top"/>
    </xsl:template>
    
    <xsl:template match="Enum" mode="top">
<xsl:if test="string-length(Help)">
#cs
<xsl:value-of select="Help"/>
#ce</xsl:if><xsl:text xml:space="preserve">
</xsl:text><xsl:value-of select="$scope"/> Enum _ ; <xsl:value-of select="@name"/><xsl:apply-templates select="Properties" mode="list"/>
    </xsl:template>
    
    <xsl:template match="Properties" mode="list">
        <xsl:if test="0 &lt; count(*)">
            <xsl:for-each select="*"><xsl:text xml:space="preserve">
</xsl:text><xsl:value-of select="$indent"/>$<xsl:value-of select="@name"/> = <xsl:value-of select="Value"/><xsl:if test="position()!=last()">, _</xsl:if>
            </xsl:for-each>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="*"/>
</xsl:stylesheet>
