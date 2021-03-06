<?xml version="1.0" encoding="UTF-8"?>
<!-- ***************************************************************************
 $Revision: 1.3 $
 $Date: 2010/07/29 14:29:00 $
 Author: doudou
 Description: Template for HTML exporter (multiple files).
 *************************************************************************** -->
<xsl:stylesheet version="2.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:fn="http://www.w3.org/2005/xpath-functions"
    xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
    xmlns:p="http://autoitscript/tldoc/xslt">
    
    <!-- functions -->
    <xsl:include href="tldoc-html-functions.inc.xsl"/>
    <!-- common templates -->
    <xsl:include href="tldoc-html-common.inc.xsl"/>
    
    <xsl:output
        omit-xml-declaration="yes"
        encoding="UTF-8"
        indent="yes"
        method="html"
        media-type="text/html"
        exclude-result-prefixes="xsl xs fn xdt p"/>
    
    <xsl:param name="key"/>
    <xsl:param name="keyVal"/>
    <xsl:param name="respath">./</xsl:param>
    <xsl:param name="definitions" select="document('tldoc-defs.inc.xml')"/>

    <xsl:key name="type" match="TypeLib | Types/*" use="concat(local-name(), '-', @name)"/>
    <xsl:key name="member" match="Properties/* | Methods/*" use="concat(local-name(../..), '-', ../../@name, '-', local-name(), '-', @name)"/>
    
    <!-- entry point -->
    <xsl:template match="/">
        <xsl:variable name="target" select="key($key, $keyVal)"/>
<html>
    <head>
        <title>Type Library Inspector: <xsl:value-of select="/TypeLib/@name"/> - <xsl:value-of select="concat(local-name($target), ' ', $target/@name)"/></title>
        <link rel="stylesheet" type="text/css" href="{$respath}css/tldoc-view.css" />
        <link rel="stylesheet" type="text/css" href="{$respath}css/tldoc-html-multi.css" />
    </head>
    <body>
        <xsl:apply-templates select="$target" mode="breadcrumbs"/>
        <xsl:choose>
            <xsl:when test="0 &lt; string-length($key) and 0 &lt; count($target)"><xsl:apply-templates select="$target" mode="top"/></xsl:when>
            <xsl:otherwise><h3>[Invalid Selection]</h3></xsl:otherwise>
        </xsl:choose>
        <p class="foot">This document was generated by Type Library Inspector <xsl:text disable-output-escaping="yes">&amp;copy;</xsl:text> 2010</p>
    </body>
</html>
    </xsl:template>
    
    <xsl:template match="Help" mode="helplink">
        <xsl:if test="string-length(@file) &gt; 0">
            <a class="icoleft onlinehelp" title="View context help for this element" href="javascript:showHelp('{translate(@file, '\', '/')}', {@context})">Online Help</a>
        </xsl:if>
    </xsl:template>

    <xsl:template match="*[@name]" mode="breadcrumbs">
        <p class="breadcrumbs">
            <xsl:if test="local-name()!='TypeLib'"><xsl:apply-templates select="/TypeLib" mode="linkedidentifier"/></xsl:if>
            <xsl:if test="local-name(..)='Properties' or local-name(..)='Methods'">
                <xsl:text> - </xsl:text><xsl:apply-templates select="../.." mode="linkedidentifier"/>
            </xsl:if>
            <xsl:if test="local-name()!='TypeLib'"><xsl:text> - </xsl:text></xsl:if><xsl:call-template name="identifier"/>
        </p>
    </xsl:template>
    
    <xsl:template match="*[@name]" mode="linkedidentifier">
        <xsl:choose>
            <xsl:when test="local-name()='TypeLib' or local-name(..)='Types'"><a href="{local-name()}-{@name}.html"><xsl:call-template name="identifier"/></a></xsl:when>
            <xsl:when test="local-name(..)='Methods' or local-name(..)='Properties'"><a href="{local-name(../..)}-{../../@name}-{local-name()}-{@name}.html"><xsl:call-template name="identifier"/></a></xsl:when>
            <xsl:otherwise><xsl:call-template name="identifier"/></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template match="*"/>
</xsl:stylesheet>
