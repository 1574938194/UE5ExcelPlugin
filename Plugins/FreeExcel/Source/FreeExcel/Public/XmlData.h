/**
*        Copyright(c) 2022  Weijian Tian
*
*    Permission is hereby granted, free of charge, to any person obtaining a copy
*    of this softwareand associated documentation files(the "Software"), to
*    deal in the Software without restriction, including without limitation the
*    rights to use, copy, modify, merge, publish, distribute, sublicense, and /or
*    sell copies of the Software, and to permit persons to whom the Software is
*    furnished to do so, subject to the following conditions :
*
*    The above copyright noticeand this permission notice shall be included in
*    all copies or substantial portions of the Software.
*
*   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
*    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
*    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
*    IN THE SOFTWARE.
*/

#pragma once

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

#include <cstring>
#include <memory>
#include <sstream>
#include <string>
#include <functional>
#include <map>
#include <vector>
#include <cstring>
#include <algorithm>
 
#include <pugixml.hpp>
using XMLNode = pugi::xml_node;
using XMLAttribute = pugi::xml_attribute;
using XMLDocument = pugi::xml_document;
  
 

enum class ContentType {
    Workbook,
    WorkbookMacroEnabled,
    Worksheet,
    Chartsheet,
    ExternalLink,
    Theme,
    Styles,
    SharedStrings,
    Drawing,
    Chart,
    ChartStyle,
    ChartColorStyle,
    ControlProperties,
    CalculationChain,
    VBAProject,
    CoreProperties,
    ExtendedProperties,
    CustomProperties,
    Comments,
    Table,
    VMLDrawing,
    Unknown
};



class  _XML_DATA
{
public:

    _XML_DATA() = default;
    _XML_DATA(const _XML_DATA& other) = delete;
    _XML_DATA& operator=(const _XML_DATA& other) = delete;
   
    _XML_DATA(std::function<std::string(std::string)> loader, const std::string& xmlPath, const std::string& xmlId= "", ContentType xmlType = ContentType::Unknown)
        :lazyLoader(loader),
        path(xmlPath),
        rid(xmlId),
        type(xmlType),
        xmlDoc(std::make_unique<XMLDocument>())
    {
        xmlDoc->reset();
    }

    void setRawData(const std::string& data)
    {
        xmlDoc->load_string(data.c_str(), pugi::parse_default | pugi::parse_ws_pcdata);
    }

    std::string getRawData() const
    {
        std::ostringstream ostr;
        getXmlDocument()->save(ostr, "", pugi::format_raw);
        return ostr.str();
    }
  
    std::string getXmlPath() const
    {
        return path;
    }

    std::string getXmlID() const
    {
        return rid;
    }

    ContentType getXmlType() const
    {
        return type;
    }

    XMLDocument* getXmlDocument()
    {
        if (!xmlDoc->document_element())
            xmlDoc->load_string(lazyLoader(path).c_str(), pugi::parse_default | pugi::parse_ws_pcdata);

        return xmlDoc.get();
    }

    const XMLDocument* getXmlDocument() const
    {
        if (!xmlDoc->document_element())
            xmlDoc->load_string(lazyLoader(path).c_str(), pugi::parse_default | pugi::parse_ws_pcdata);

        return xmlDoc.get();
    }

    void setXmlData(const std::string& xmlData)
    {
        setRawData(xmlData);
    }


    std::string xmlData() const
    {
        return getRawData();
    }
 
    std::string relationshipID() const
    {
        return getXmlID();
    }

    XMLDocument& xmlDocument()
    {
        return *getXmlDocument();    // NOLINT
    }

    const XMLDocument& xmlDocument() const
    {
        return *getXmlDocument();
    }

private: 
    std::function<std::string(std::string)> lazyLoader;
    std::string path{};   /**< The path of the XML data in the .xlsx zip archive. >*/
    std::string rid{};     /**< The relationship ID of the XML data. >*/
    ContentType type{};   /**< The type represented by the XML data. >*/
    std::unique_ptr<XMLDocument> xmlDoc;       /**< The underlying XMLDocument object. >*/
};



#pragma warning(pop)
 