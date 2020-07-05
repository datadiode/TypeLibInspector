<?xml version="1.0" encoding="UTF-8"?>
<!-- ***************************************************************************
 $Revision: 1.3 $
 $Date: 2010/07/29 14:25:09 $
 Author: doudou
 Description: Template for TypeLibInspector's main view.
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
        indent="yes"
        method="html"
        media-type="text/html"
        exclude-result-prefixes="xsl xs fn xdt p"/>
    
    <xsl:param name="ptr">16</xsl:param>
    <xsl:param name="respath">../</xsl:param>
    <xsl:param name="definitions" select="document('tldoc-defs.inc.xml')"/>
        
    <!-- functions -->
    <xsl:include href="tldoc-html-functions.inc.xsl"/>
    <!-- common templates -->
    <xsl:include href="tldoc-html-common.inc.xsl"/>
    
    <!-- entry point -->
    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="0 &lt; string-length($ptr) and 0 &lt; count(//*[@tldapp-ptr=$ptr])"><xsl:apply-templates select="//*[@tldapp-ptr=$ptr]" mode="top"/></xsl:when>
            <xsl:otherwise><h3>[Invalid Selection]</h3></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template match="Help" mode="helplink">
        <xsl:if test="string-length(@file) &gt; 0">
            <a class="icoleft onlinehelp" title="View context help for this element" href="?cmd=showHelp&amp;file={@file}&amp;context={@context}">Online Help</a>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="*[@name]" mode="linkedidentifier">
        <xsl:choose>
            <xsl:when test="string-length(@tldapp-ptr) &gt; 0"><a href="?cmd=showNode&amp;ptr={@tldapp-ptr}"><xsl:call-template name="identifier"/></a></xsl:when>
            <xsl:when test="string-length(@typelib) &gt; 0"><a href="?cmd=showExternal&amp;typelib={@typelib}&amp;name={@name}&amp;guid={@guid}&amp;number={@number}"><xsl:call-template name="identifier"/></a></xsl:when>
            <xsl:otherwise><xsl:call-template name="identifier"/></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template match="*"/>
</xsl:stylesheet>
