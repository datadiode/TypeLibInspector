<?xml version="1.0" encoding="UTF-8"?>
<!-- ***************************************************************************
 $Revision: 1.2 $
 $Date: 2010/07/23 16:56:53 $
 Author: doudou
 Description: Common templates (functions) for HTML transformation.
 *************************************************************************** -->
<xsl:stylesheet version="2.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:fn="http://www.w3.org/2005/xpath-functions"
    xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
    xmlns:p="http://autoitscript/tldoc/xslt">
    
    <xsl:template name="infotophead">
        <h2 id="{generate-id()}"><xsl:value-of select="local-name()"/></h2>
        <h1><xsl:attribute name="class">icolefttop 
        <xsl:choose>
            <xsl:when test="local-name()='Property' and Attributes/Flag[@val='1']"><xsl:value-of select="local-name()"/>RO</xsl:when>
            <xsl:when test="local-name()='Property' and @kind='2'">Const</xsl:when>
            <xsl:otherwise><xsl:value-of select="local-name()"/></xsl:otherwise>
        </xsl:choose>
        </xsl:attribute><xsl:value-of select="@name"/></h1>
        <hr/>
        <p><xsl:value-of select="Help"/></p>
    </xsl:template>
    
    <xsl:template name="typeinfohead">
        <xsl:param name="flags" select="$definitions/*/p:typeflags"/>
        <div><strong>GUID: </strong>
        <xsl:value-of select="@guid"/></div>
        <div><strong>Version: </strong>
        <xsl:value-of select="Version/@major"/>.<xsl:value-of select="Version/@minor"/></div>
        <xsl:if test="Attributes/@mask &gt; 0">
            <h4>Attributes</h4>
            <dl>
                <xsl:for-each select="Attributes/Flag">
                    <xsl:variable name="val" select="@val"/>
                    <xsl:variable name="fn" select="$flags/p:flag[@val=$val]"/>
                    <dt><xsl:value-of select="$fn/@name"/></dt>
                    <dd><xsl:value-of select="$fn"/></dd>
                </xsl:for-each>
            </dl>
        </xsl:if>
    </xsl:template>
    
    <xsl:template name="memberflags">
        <xsl:param name="flags" select="$definitions/*/p:funcflags"/>
        <xsl:if test="Attributes/@mask &gt; 0">
            <h4>Attributes</h4>
            <dl>
                <xsl:for-each select="Attributes/Flag">
                    <xsl:variable name="val" select="@val"/>
                    <xsl:variable name="fn" select="$flags/p:flag[@val=$val]"/>
                    <dt><xsl:value-of select="$fn/@name"/></dt>
                    <dd><xsl:value-of select="$fn"/></dd>
                </xsl:for-each>
            </dl>
        </xsl:if>
    </xsl:template>
    
    <xsl:template name="identifier">
        <xsl:choose>
            <xsl:when test="local-name()='type' or local-name()='TypeLib' or local-name()='TypeRef' or local-name(..)='Types'"><dfn class="typename" title="{local-name()} {@name}"><xsl:value-of select="@name"/></dfn></xsl:when>
            <xsl:otherwise><span class="identifier" title="{local-name()} {@name}"><xsl:value-of select="@name"/></span></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template name="value">
        <xsl:param name="val"/>
        <xsl:param name="vt"/>
        <xsl:variable name="isstr" select="boolean($vt/@indirection &lt; 1 and ($vt/@vt='8' or $vt/@vt='30' or $vt/@vt='31'))"/>
        <xsl:if test="$isstr">&quot;</xsl:if><xsl:value-of select="$val"/><xsl:if test="$isstr">&quot;</xsl:if>
    </xsl:template>
</xsl:stylesheet>
