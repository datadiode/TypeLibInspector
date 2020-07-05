<?xml version="1.0" encoding="UTF-8"?>
<!-- ***************************************************************************
 $Revision: 1.4 $
 $Date: 2010/07/29 14:29:00 $
 Author: doudou
 Description: Common templates for HTML transformation.
 *************************************************************************** -->
<xsl:stylesheet version="2.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:fn="http://www.w3.org/2005/xpath-functions"
    xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
    xmlns:p="http://autoitscript/tldoc/xslt">

    <xsl:template match="TypeLib" mode="top">
        <xsl:call-template name="infotophead"/>
        <div><strong>LCID: </strong> <xsl:value-of select="@lcid"/></div>
        <div><strong>Target system: </strong>
        <xsl:choose>
            <xsl:when test="@sysKind=0">Win16</xsl:when>
            <xsl:when test="@sysKind=1">Win32</xsl:when>
            <xsl:when test="@sysKind=2">Mac</xsl:when>
            <xsl:when test="@sysKind=3">Win64</xsl:when>
            <xsl:otherwise>Unknown</xsl:otherwise>
        </xsl:choose></div>
        <div><strong>Source: </strong>
        <xsl:value-of select="@path"/></div>
        <xsl:call-template name="typeinfohead">
            <xsl:with-param name="flags" select="$definitions/*/p:libflags"/>
        </xsl:call-template>
        <xsl:apply-templates select="Types" mode="list"/>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="Types" mode="list">
        <xsl:if test="0 &lt; count(Enum)">
            <h3>Enums</h3>
            <ul class="memblist">
            <xsl:apply-templates select="Enum" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(Record)">
            <h3>Records</h3>
            <ul class="memblist">
            <xsl:apply-templates select="Record" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(Interface)">
            <h3>Interfaces</h3>
            <ul class="memblist">
            <xsl:apply-templates select="Interface" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(DispInterface)">
            <h3>DispInterfaces</h3>
            <ul class="memblist">
            <xsl:apply-templates select="DispInterface" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(CoClass)">
            <h3>CoClasses</h3>
            <ul class="memblist">
            <xsl:apply-templates select="CoClass" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(Alias)">
            <h3>Aliases</h3>
            <ul class="memblist">
            <xsl:apply-templates select="Alias" mode="list"/>
            </ul>
        </xsl:if>
    </xsl:template>

    <xsl:template match="Types/*" mode="list">
        <li class="icoleft {local-name()}"><span class="h16"></span><xsl:apply-templates select="." mode="linkedidentifier"/></li>
    </xsl:template>
    
    <xsl:template match="DispInterface | Interface | Record | Module | Enum" mode="top">
        <xsl:call-template name="infotophead"/>
        <div><strong>Defined in: </strong>
        <xsl:apply-templates select="/TypeLib" mode="linkedidentifier"/></div>
        <xsl:call-template name="typeinfohead"/>
        <hr/>
        <xsl:choose>
            <xsl:when test="local-name()='Interface' or local-name()='DispInterface'">
                <xsl:if test="0 &lt; count(Base/*)">
                    <h3>Base Interface</h3>
                    <ul class="memblist">
                        <xsl:apply-templates select="Base" mode="list"/>
                    </ul>
                </xsl:if>
            </xsl:when>
            <xsl:when test="local-name()='DispInterface'">
                <xsl:if test="0 &lt; count(VTable/*)">
                    <h3>VTable Interface</h3>
                    <ul class="memblist">
                        <xsl:apply-templates select="VTable" mode="list"/>
                    </ul>
                </xsl:if>
            </xsl:when>
        </xsl:choose>
        <xsl:apply-templates select="Properties" mode="list"/>
        <xsl:apply-templates select="Methods" mode="list"/>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="CoClass" mode="top">
        <xsl:call-template name="infotophead"/>
        <div><strong>Defined in: </strong>
        <xsl:apply-templates select="/TypeLib" mode="linkedidentifier"/></div>
        <xsl:if test="@progid and @progid != '0'">
            <div><strong>ProgID: </strong>
            <xsl:value-of select="@progid"/></div>
        </xsl:if>
        <xsl:call-template name="typeinfohead"/>
        <hr/>
        <h3>Interfaces</h3>
        <ul class="memblist">
        <xsl:apply-templates select="Interfaces/*" mode="list"/>
        </ul>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="Alias" mode="top">
        <xsl:call-template name="infotophead"/>
        <div><strong>Defined in: </strong>
        <xsl:apply-templates select="/TypeLib" mode="linkedidentifier"/></div>
        <xsl:call-template name="typeinfohead"/>
        <hr/>
        <h3>Resolves To</h3>
        <xsl:variable name="name" select="Resolved/VarType/TypeRef/@name"/>
        <xsl:variable name="guid" select="Resolved/VarType/TypeRef/@guid"/>
        <xsl:variable name="origin" select="/TypeLib/Types/*[@name=$name and @guid=$guid]"/>
        <ul class="memblist">
        <li class="icoleft {local-name($origin)}"><span class="h16"></span><xsl:apply-templates select="Resolved/VarType" mode="any"/></li>
        </ul>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="Properties" mode="list">
        <xsl:if test="0 &lt; count(*)">
            <h3>Properties</h3>
            <ul class="memblist">
            <xsl:apply-templates select="*" mode="list"/>
            </ul>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="Methods" mode="list">
        <xsl:if test="0 &lt; count(PropertyGet | PropertyPut | PropertyPutRef)">
            <h3>Properties (Indirect)</h3>
            <ul class="memblist">
            <xsl:apply-templates select="PropertyGet | PropertyPut | PropertyPutRef" mode="list"/>
            </ul>
        </xsl:if>
        <xsl:if test="0 &lt; count(Function | Method)">
            <h3>Methods</h3>
            <ul class="memblist">
            <xsl:apply-templates select="Function | Method" mode="list"/>
            </ul>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="Impl | VTable | Base" mode="list">
        <li class="icoleft {local-name()}"><span class="h16"></span>
            <xsl:choose>
                <xsl:when test="TypeRef"><xsl:apply-templates select="TypeRef" mode="resolveidentifier"/></xsl:when>
                <xsl:otherwise><xsl:apply-templates select="*[position() = 1]" mode="linkedidentifier"/></xsl:otherwise>
            </xsl:choose>
            <xsl:if test="local-name()='Impl'">
                <xsl:if test="Attributes/Flag/@val='1'"> <img width="16" height="16" src="{$respath}icons/default.ico" align="middle" alt="default" title="Default (represents the default for the source or sink)"/></xsl:if>
                <xsl:if test="Attributes/Flag/@val='2'"> <img width="16" height="16" src="{$respath}icons/source.ico" align="middle" alt="source" title="Source (this member is called rather than implemented)"/></xsl:if>
            </xsl:if>
        </li>
    </xsl:template>
    
    <xsl:template match="Property" mode="list">
        <li><xsl:attribute name="class">icoleft 
        <xsl:choose>
            <xsl:when test="Attributes/Flag[@val='1']"><xsl:value-of select="local-name()"/>RO</xsl:when>
            <xsl:when test="@kind='2'">Const</xsl:when>
            <xsl:otherwise><xsl:value-of select="local-name()"/></xsl:otherwise>
        </xsl:choose>
        </xsl:attribute><span class="h16"></span><code><xsl:apply-templates select="." mode="linkedidentifier"/>
        <xsl:if test="Value"><xsl:text> = </xsl:text><xsl:value-of select="Value"/></xsl:if></code>
        <xsl:if test="Attributes/Flag/@val='32'"> <img src="{$respath}icons/default.ico" align="middle" alt="defaultbind" title="Default"/></xsl:if>
        </li>
    </xsl:template>
    
    <xsl:template match="Property" mode="top">
        <xsl:call-template name="infotophead"/>
        <code class="membsignature"><xsl:choose>
            <xsl:when test="@kind='1'">static type</xsl:when>
            <xsl:when test="@kind='2'">const type</xsl:when>
            <xsl:otherwise>object</xsl:otherwise>
        </xsl:choose>.<xsl:value-of select="@name"/><xsl:if test="Value"><xsl:text> = </xsl:text><xsl:call-template name="value"><xsl:with-param name="val" select="Value"/><xsl:with-param name="vt" select="VarType/@vt"/></xsl:call-template></xsl:if></code>
        <xsl:call-template name="memberflags"><xsl:with-param name="flags" select="$definitions/*/p:varflags"/></xsl:call-template>
        <xsl:choose>
            <xsl:when test="@kind='1' or @kind='2'">
                <h3>Defined in</h3>
                <dl class="subject">
                    <dt>type</dt>
                    <dd><xsl:apply-templates select="../.." mode="linkedidentifier"/></dd>
                </dl>
            </xsl:when>
            <xsl:otherwise>
                <h3>Applies To</h3>
                <dl class="subject">
                    <dt>object</dt>
                    <dd>Instance of <xsl:apply-templates select="../.." mode="linkedidentifier"/></dd>
                </dl>
            </xsl:otherwise>
        </xsl:choose>
        <h3>Value</h3>
        <dl class="returnval">
            <dt>
            <xsl:choose>
                <xsl:when test="Value"><xsl:call-template name="value"><xsl:with-param name="val" select="Value"/><xsl:with-param name="vt" select="VarType/@vt"/></xsl:call-template></xsl:when>
                <xsl:otherwise><xsl:value-of select="@name"/> takes values of </xsl:otherwise>
            </xsl:choose>
            </dt>
            <dd><xsl:apply-templates select="VarType" mode="any"/></dd>
        </dl>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="PropertyGet | PropertyPut | PropertyPutRef | Function" mode="list">
        <li class="icoleft {local-name()}"><span class="h16"></span>
        <code><xsl:apply-templates select="." mode="linkedidentifier"/>(<xsl:apply-templates select="Parameters" mode="paramlist"/>)</code>
        <xsl:if test="Attributes/Flag/@val='32'"> <img width="16" height="16" src="{$respath}icons/default.ico" align="middle" alt="defaultbind" title="Default"/></xsl:if>
        </li>
    </xsl:template>
    
    <xsl:template match="PropertyGet | PropertyPut | PropertyPutRef | Function" mode="top">
        <xsl:call-template name="infotophead"/>
        <code class="membsignature">object.<xsl:value-of select="@name"/>(<xsl:apply-templates select="Parameters" mode="paramlist"/>)</code>
        <xsl:call-template name="memberflags"/>
        <h3>Applies To</h3>
        <dl class="subject">
            <dt>object</dt>
            <dd>Instance of <xsl:apply-templates select="../.." mode="linkedidentifier"/></dd>
        </dl>
        <h3>Parameters</h3>
        <xsl:apply-templates select="Parameters" mode="paramdesc"/>
        <h3>Return Value</h3>
        <dl class="returnval">
            <dt><xsl:value-of select="@name"/> returns</dt>
            <dd><xsl:apply-templates select="VarType" mode="any"/></dd>
        </dl>
        <hr/>
        <xsl:apply-templates select="Help" mode="helplink"/>
    </xsl:template>
    
    <xsl:template match="Parameters" mode="paramdesc">
        <xsl:choose>
            <xsl:when test="count(Parameter)">
        <dl class="paramdesc">
            <xsl:for-each select="Parameter">
                <dt><xsl:value-of select="@name"/></dt>
                <dd>
                    <xsl:if test="@optional='1'">
                        <strong>[optional] </strong>
                    </xsl:if>
                    <xsl:if test="Flags/Flag/@val='2'">
                        <strong>[out] </strong>
                    </xsl:if>
                    <xsl:if test="Flags/Flag/@val='8'">
                        <strong>[retval] </strong>
                    </xsl:if>
                    <xsl:apply-templates select="VarType" mode="any"/>
                </dd>
                <xsl:if test="0 &lt; string-length(Value)">
                    <dd>Default: <xsl:call-template name="value"><xsl:with-param name="val" select="Value"/><xsl:with-param name="vt" select="VarType"/></xsl:call-template></dd>
                </xsl:if>
            </xsl:for-each>
        </dl>
            </xsl:when>
            <xsl:otherwise><div>None</div></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <xsl:template match="Parameters" mode="paramlist">
        <xsl:for-each select="Parameter">
            <xsl:if test="@optional='1'">
                <xsl:text>[</xsl:text>
            </xsl:if>
            <xsl:if test="Flags/Flag/@val='2' or Flags/Flag/@val='8'">&lt;</xsl:if>
            <xsl:if test="1 &lt; position()">
                <xsl:text>, </xsl:text>
            </xsl:if>
            <xsl:value-of select="@name"/>
            <xsl:if test="Flags/Flag/@val='2'"><sup> out</sup></xsl:if>
            <xsl:if test="Flags/Flag/@val='8'"><sup> ret</sup></xsl:if>
            <xsl:if test="0 &lt; string-length(Value)"><xsl:text> = </xsl:text><xsl:call-template name="value"><xsl:with-param name="val" select="Value"/><xsl:with-param name="vt" select="VarType"/></xsl:call-template></xsl:if>
            <xsl:if test="Flags/Flag/@val='2' or Flags/Flag/@val='8'">&gt;</xsl:if>
            <xsl:if test="@optional='1'">
                <xsl:text>]</xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template match="VarType" mode="any">
        <xsl:if test="@indirection &gt; 0">Pointer to </xsl:if>
        <xsl:if test="@indirection &gt; 1">pointer to </xsl:if>
        <xsl:if test="@indirection &gt; 2">pointer to </xsl:if>
        <xsl:variable name="vt" select="@vt"/>
        <xsl:variable name="isptr" select="boolean(@indirection &gt; 0)"/>
        <xsl:variable name="tn" select="$definitions/*/p:typedmap/p:type[@vt=$vt and (boolean(@byref)=$isptr or string(@byref)='')][position() = last()]"/>
        <xsl:choose>
            <xsl:when test="TypeRef"><xsl:apply-templates select="TypeRef" mode="resolveidentifier"/></xsl:when>
            <xsl:otherwise><xsl:apply-templates select="$tn" mode="linkedidentifier"/></xsl:otherwise>
        </xsl:choose>
        <xsl:text> (</xsl:text>
        <xsl:value-of select="$tn"/>
        <xsl:text>)</xsl:text>
        <xsl:if test="Array"> of
            <xsl:for-each select="Array/Dim">
                <xsl:if test="1 &lt; position()"><xsl:text disable-output-escaping="yes"> &amp;times; </xsl:text></xsl:if>
                [<xsl:value-of select="@lbound"/> - <xsl:value-of select="@ubound"/>]
            </xsl:for-each>
            <xsl:if test="0 &lt; count(Array/Dim)"><xsl:text disable-output-escaping="yes"> &amp;times; </xsl:text></xsl:if>
        <xsl:apply-templates select="Array/VarType" mode="any"/></xsl:if>
    </xsl:template>
    
    <xsl:template match="TypeRef" mode="top">
        <h2 id="{generate-id()}">Type Reference</h2>
        <h1><xsl:value-of select="@name"/></h1>
        <hr/>
        <div><strong>GUID: </strong>
        <xsl:value-of select="@guid"/></div>
        <div><strong>TypeLib GUID: </strong>
        <xsl:value-of select="@typelib"/></div>
    </xsl:template>

    <xsl:template match="TypeRef" mode="resolveidentifier">
        <xsl:variable name="name" select="@name"/>
        <xsl:variable name="guid" select="@guid"/>
        <xsl:variable name="origin" select="/TypeLib/Types/*[@name=$name and @guid=$guid]"/>
        <xsl:choose>
            <xsl:when test="$origin/@name"><xsl:apply-templates select="$origin" mode="linkedidentifier"/></xsl:when>
            <xsl:otherwise><xsl:apply-templates select="." mode="linkedidentifier"/></xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
