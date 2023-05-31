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

#include "CoreMinimal.h"
#include "UObject/NoExportTypes.h"  
#include "CellReference.h"
#include "Templates/UniquePtr.h"
#include "zippy.h"
#include "XmlData.h"
#include <filesystem>
#include <list>
#include <queue>
#include "Document.generated.h"

class USheet;
 
/**
 * 
 */
UCLASS(Blueprintable)
class FREEEXCEL_API UDocument : public UObject
{
    GENERATED_UCLASS_BODY()
public:
     
    enum class RelsType {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        Chartsheet,
        Dialogsheet,
        Macrosheet,
        CalculationChain,
        ExternalLink,
        ExternalLinkPath,
        Theme,
        Styles,
        Chart,
        ChartStyle,
        ChartColorStyle,
        Image,
        Drawing,
        VMLDrawing,
        SharedStrings,
        PrinterSettings,
        VBAProject,
        ControlProperties,
        Unknown
    };

    enum class ExcelProperty {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        LastModifiedBy,
        LastPrinted,
        CreationDate,
        ModificationDate,
        Category,
        Application,
        DocSecurity,
        ScaleCrop,
        Manager,
        Company,
        LinksUpToDate,
        SharedDoc,
        HyperlinkBase,
        HyperlinksChanged,
        AppVersion
    };
 

public:

    //  Open the .xlsx file with the given path
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void Open(FString path);

    // Create a new .xlsx file with the given path.
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void Create(FString path);

    // Close the current document
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void Close();

    //Save the current document using the current filename, overwriting the existing file.
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void Save();

    //Save the document with a new name. If a file exists with that name, it will be overwritten.
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void SaveAs(FString path);

    // Get the filename of the current document, e.g. "spreadsheet.FreeExcel".
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
        FString FileName()const;

    //Get the full path of the current document, e.g. "drive/blah/spreadsheet.FreeExcel"
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
        FString FullPath()const;

    // Check the xlsx document is valid
    UFUNCTION(BlueprintPure, meta = (DisplayName = "Is Valid (ExcelDocument)"), Category = "FreeExcel")
        bool IsValid() const;

    // Get the sheet with name, and set CurrentSheet to it
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        USheet* GetOrCreateSheetWithName(FString name);
 
    // Get last access Sheet object
    USheet* GetCurrentSheet()
    {
        return _curSheet;
    }
 
    // Get .xlsx property value
    std::string GetProperty(ExcelProperty prop) const; 
 
    // Set .xlsx property value
    void SetProperty(ExcelProperty prop, const std::string& value);
 
    // Remove .xlsx proprty 
    void RemoveProperty(ExcelProperty theProperty);

    void SetSheetName(std::string id, FString NewName);
     
    void ResetCalcChain();
     
    // Add new sheet
    void AddWorksheet(FString SheetName);

    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
    void CloneSheet(FString SheetID, FString CloneName);
   
     const _XML_DATA* QueryXmlData(std::string xmlPath) const;
     
protected:
 
    void AddSharedStrings();
    std::string extractXmlFromArchive(const std::string& path);
 
    _XML_DATA* getXmlData(const std::string& path);
 
    const _XML_DATA* getXmlData(const std::string& path) const;
 
    bool hasXmlData(const std::string& path) const;
   
protected:

#pragma region Core Properties

public:  
          
    void setCoreProperty(const std::string& name, const std::string& value);
  
    std::string getCoreProperty(const std::string& name) const;
 
    void removeCoreProperty(const std::string& name);

#pragma endregion

#pragma region App Properties

    inline XMLAttribute headingPairsSize(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().attribute("size");
    }

    inline XMLNode headingPairsCategories(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().first_child();
    }

    inline XMLNode headingPairsCounts(XMLNode docNode)
    {
        return headingPairsCategories(docNode).next_sibling();
    }

    inline XMLNode sheetNames(XMLNode docNode)
    {
        return docNode.child("TitlesOfParts").first_child();
    }

    inline XMLAttribute sheetCount(XMLNode docNode)
    {
        return sheetNames(docNode).attribute("size");
    }


    void AddSheetName(const std::string& title);

    void RemoveSheetName(const std::string& title);
     
    void AddHeadingPair(const std::string& name, int value);

    void RemoveHeadingPair(const std::string& name);

    void SetHeadingPair(const std::string& name, int newValue);

    void setAppProperty(const std::string& name, const std::string& value);
         
    std::string getAppProperty(const std::string& name) const; 

    void RemoveAppProperty(const std::string& name); 

    void AppendSheetName(const std::string& sheetName); 

    void PrependSheetName(const std::string& sheetName); 

    void InsertSheetName(const std::string& sheetName, unsigned int index);

#pragma endregion
      
public:
    int32_t  getStringIndex(const std::string& str) const
    {
        auto iter = std::find_if(sharedStringCache.begin(), sharedStringCache.end(), [&](const std::string& s) { return str == s; });

        return iter == sharedStringCache.end() ? -1 : static_cast<int32_t>(std::distance(sharedStringCache.begin(), iter));
    }

    bool  stringExists(const std::string& str) const
    {
        return getStringIndex(str) >= 0;
    }
     
    const char* getString(uint32_t index) const
    {
        return sharedStringCache[index].c_str();
    }

    int32_t  appendString(const std::string& str)
    {
        auto textNode = shared_strings->xmlDocument().document_element().append_child("si").append_child("t");
        if (str.front() == ' ' || str.back() == ' ') textNode.append_attribute("xml:space").set_value("preserve");

        textNode.text().set(str.c_str());
        sharedStringCache.emplace_back(textNode.text().get());

        return static_cast<int32_t>(std::distance(sharedStringCache.begin(), sharedStringCache.end()) - 1);
    }

    void  clearString(uint64_t index)
    {
        sharedStringCache[index] = "";
        auto iter = shared_strings->xmlDocument().document_element().children().begin();
        std::advance(iter, index);
        iter->text().set("");
    }

 
     
#pragma region Workbook
    static XMLNode sheetsNode(const XMLDocument& doc) 
    {
        return doc.document_element().child("sheets");
    }

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    USheet* GetSheet(FString name) const;
     
    USheet* GetSheetAt(uint16_t index);
 
    bool hasSharedStrings() const
    {
        return true;//parentDoc().executeQuery(XLQuerySharedStrings()).sharedStrings() != nullptr;
    }
     

    void deleteNamedRanges();

    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        void  RemoveSheet(FString name);
 
    uint16_t createInternalSheetID();

    std::string sheetID(const std::string& sheetName)
    {
        return workbook->xmlDocument().document_element().child("sheets").find_child_by_attribute("name", sheetName.c_str()).
            attribute("r:id").value();
    }

    std::string sheetName(const std::string& sheetID) const
    {
        return workbook->xmlDocument().document_element().child("sheets").find_child_by_attribute("r:id", sheetID.c_str()).
            attribute("name").value();
    }


    void prepareSheetMetadata(const std::string& sheetName, uint16_t internalID);
 
    inline  XMLNode findSheetWithName(std::string sheetName)
    {
        return workbook->xmlDocument().
            find_child_by_attribute("name", sheetName.c_str());
    }
    inline const XMLNode findSheetWithName(std::string sheetName) const
    {
        return workbook->xmlDocument().
            find_child_by_attribute("name", sheetName.c_str());
    }
    inline  XMLNode findSheetWithRID(std::string sheetID)
    {
        return workbook->xmlDocument().document_element().child("sheets").
            find_child_by_attribute("r:id", sheetID.c_str());
    }
    inline const XMLNode findSheetWithRID(std::string sheetID) const
    {
        return workbook->xmlDocument().document_element().child("sheets").
            find_child_by_attribute("r:id", sheetID.c_str());
    }
    bool GetSheetVisible(const std::string& name) const;

    void SetSheetVisible(const std::string& id, bool state);

    void setSheetIndex(const std::string& sheetName, int index);
     
    int32 IndexOfSheet(const std::string& sheetName) const;

     int sheetCount() const
    {
        return std::distance(sheetsNode(workbook->xmlDocument()).children().begin(),
            sheetsNode(workbook->xmlDocument()).children().end());
    }

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    int32 GetSheetCount() const
    {
        return worksheetNames().size();
    }
 
    //std::vector<std::string> sheetNames() const
    //{
    //    std::vector<std::string> results;

    //    for (const auto& item : sheetsNode(workbook->xmlDocument()).children()) results.emplace_back(item.attribute("name").value());

    //    return results;
    //}
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        TArray<FString> GetSheetNames() const;

    std::vector<std::string> worksheetNames() const;

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    bool SheetExists(FString name) const
    {
        std::string sheetName = TCHAR_TO_UTF8(*name);
        return  worksheetExists(sheetName);
    }


    bool worksheetExists(const std::string& sheetName) const
    {
        auto wksNames = worksheetNames();
        return std::find(wksNames.begin(), wksNames.end(), sheetName) != wksNames.end();
    }
  
    // The UpdateSheetName member function searches throug the usages of the old name and replaces with the new sheet name.
    void UpdateSheetReferences(const std::string& oldName, const std::string& newName);

    void SetFullCalculationOnLoad();

    bool SheetIsActive(const std::string& sheetRID) const;

    void SetSheetActive(const std::string& sheetRID);
#pragma endregion
     
#pragma region Rels
 
    std::string  docRelsIdByTarget(const std::string& target) const
    {
        return doc_rels->xmlDocument().
            document_element().
            find_child_by_attribute("Target", target.c_str()).
            attribute("Id").value();
    }

    // sheet ID by Target
    std::string  wbkRelsIdByTarget(const std::string& target) const
    {
        return workbook_rels->xmlDocument().
            document_element().
            find_child_by_attribute("Target", target.c_str()).
            attribute("Id").value();
    }

    // sheet Target by ID
    std::string wbkRelsTargetById(const std::string& id) const
    {
        return workbook_rels->xmlDocument().document_element().find_child_by_attribute("Id", id.c_str()).
            attribute("Target").value();
    }

    RelsType wbkRelsTypeById(const std::string& id) const
    {
        const std::string typeString = workbook_rels->xmlDocument().document_element().find_child_by_attribute("Id", id.c_str()).
            attribute("Type").value(); 
        RelsType type;

        if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
            type = RelsType::ExtendedProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
            type = RelsType::CustomProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
            type = RelsType::Workbook;
        else if (typeString == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
            type = RelsType::CoreProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
            type = RelsType::Worksheet;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
            type = RelsType::Styles;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
            type = RelsType::SharedStrings;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
            type = RelsType::CalculationChain;
        else if (typeString == "http://schemas.microsoft.com/office/2006/relationships/vbaProject")
            type = RelsType::VBAProject;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink")
            type = RelsType::ExternalLink;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
            type = RelsType::Theme;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
            type = RelsType::Chartsheet;
        else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartStyle")
            type = RelsType::ChartStyle;
        else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle")
            type = RelsType::ChartColorStyle;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
            type = RelsType::Drawing;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            type = RelsType::Image;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
            type = RelsType::Chart;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath")
            type = RelsType::ExternalLinkPath;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
            type = RelsType::PrinterSettings;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing")
            type = RelsType::VMLDrawing;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp")
            type = RelsType::ControlProperties;
        else
            type = RelsType::Unknown;

        return type;
    }
     

    //std::vector<_RelsItem>  relationships() const
    //{
    //    auto result = std::vector<_RelsItem>();
    //    for (const auto& item : data_->xmlDocument().document_element().children()) result.emplace_back(_RelsItem(item));

    //    return result;
    //}

    void  RemoveRelationship(const std::string& relID)
    {
        workbook_rels->xmlDocument().document_element().remove_child(
            workbook_rels->xmlDocument().document_element().find_child_by_attribute("Id", relID.c_str()));
    }
     
    static uint32_t GetNewRelsID(XMLNode relationshipsNode)
    {
        return static_cast<uint32_t>(stoi(std::string(std::max_element(relationshipsNode.children().begin(),
            relationshipsNode.children().end(),
            [](XMLNode a, XMLNode b) {
                return stoi(std::string(a.attribute("Id").value()).substr(3)) <
                    stoi(std::string(b.attribute("Id").value()).substr(3));
            })
            ->attribute("Id")
                .value())
            .substr(3)) +
            1);
    }

    void  addRelationship(RelsType type, const std::string& target)
    {
        std::string typeString = [](RelsType type)->std::string
        {
            std::string typeString;

            if (type == RelsType::ExtendedProperties)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
            else if (type == RelsType::CustomProperties)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
            else if (type == RelsType::Workbook)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            else if (type == RelsType::CoreProperties)
                typeString = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
            else if (type == RelsType::Worksheet)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            else if (type == RelsType::Styles)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            else if (type == RelsType::SharedStrings)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
            else if (type == RelsType::CalculationChain)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
            else if (type == RelsType::VBAProject)
                typeString = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
            else if (type == RelsType::ExternalLink)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
            else if (type == RelsType::Theme)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
            else if (type == RelsType::Chartsheet)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
            else if (type == RelsType::ChartStyle)
                typeString = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
            else if (type == RelsType::ChartColorStyle)
                typeString = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
            else if (type == RelsType::Drawing)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
            else if (type == RelsType::Image)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
            else if (type == RelsType::Chart)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
            else if (type == RelsType::ExternalLinkPath)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
            else if (type == RelsType::PrinterSettings)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
            else if (type == RelsType::VMLDrawing)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
            else if (type == RelsType::ControlProperties)
                typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
            else
                typeString = "unknow";

            return typeString;
        }(type);

        std::string id = "rId" + std::to_string(GetNewRelsID(workbook_rels->xmlDocument().document_element()));

        // Create new node in the .rels file
        auto node = workbook_rels->xmlDocument().document_element().append_child("Relationship");
        node.append_attribute("Id").set_value(id.c_str());
        node.append_attribute("Type").set_value(typeString.c_str());
        node.append_attribute("Target").set_value(target.c_str());

        if (type == RelsType::ExternalLinkPath) {
            node.append_attribute("TargetMode").set_value("External");
        }
         
    }

    //bool  targetExists(const std::string& target) const
    //{
    //    return data_->xmlDocument().document_element().find_child_by_attribute("Target", target.c_str()) != nullptr;
    //}


    //bool idExists(const std::string& id) const
    //{
    //    return data_->xmlDocument().document_element().find_child_by_attribute("Id", id.c_str()) != nullptr;
    //}

  
#pragma endregion

#pragma region ContentTypes
     
    ContentType ContentTypeType(XMLNode& node) const
    {
        ContentType type;
        const std::string typeString = node.attribute("ContentType").value();

        if (typeString == "application/vnd.ms-excel.Sheet.macroEnabled.main+xml")
            type = ContentType::WorkbookMacroEnabled;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            type = ContentType::Workbook;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            type = ContentType::Worksheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")
            type = ContentType::Chartsheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
            type = ContentType::ExternalLink;
        else if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml")
            type = ContentType::Theme;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
            type = ContentType::Styles;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
            type = ContentType::SharedStrings;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml")
            type = ContentType::Drawing;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            type = ContentType::Chart;
        else if (typeString == "application/vnd.ms-office.chartstyle+xml")
            type = ContentType::ChartStyle;
        else if (typeString == "application/vnd.ms-office.chartcolorstyle+xml")
            type = ContentType::ChartColorStyle;
        else if (typeString == "application/vnd.ms-excel.controlproperties+xml")
            type = ContentType::ControlProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
            type = ContentType::CalculationChain;
        else if (typeString == "application/vnd.ms-office.vbaProject")
            type = ContentType::VBAProject;
        else if (typeString == "application/vnd.openxmlformats-package.core-properties+xml")
            type = ContentType::CoreProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml")
            type = ContentType::ExtendedProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            type = ContentType::CustomProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            type = ContentType::Comments;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
            type = ContentType::Table;
        else if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing")
            type = ContentType::VMLDrawing;
        else
            type = ContentType::Unknown;

        return type;

    }
     
    std::string ContentTypePath(XMLNode& node)
    {
        return node.attribute("PartName").value();
    }

    void addOverride(const std::string& path, ContentType type)
    {
        std::string typeString = [](ContentType type)->std::string
        {
            std::string typeString;

            if (type == ContentType::WorkbookMacroEnabled)
                typeString = "application/vnd.ms-excel.Sheet.macroEnabled.main+xml";
            else if (type == ContentType::Workbook)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
            else if (type == ContentType::Worksheet)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
            else if (type == ContentType::Chartsheet)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
            else if (type == ContentType::ExternalLink)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
            else if (type == ContentType::Theme)
                typeString = "application/vnd.openxmlformats-officedocument.theme+xml";
            else if (type == ContentType::Styles)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
            else if (type == ContentType::SharedStrings)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
            else if (type == ContentType::Drawing)
                typeString = "application/vnd.openxmlformats-officedocument.drawing+xml";
            else if (type == ContentType::Chart)
                typeString = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
            else if (type == ContentType::ChartStyle)
                typeString = "application/vnd.ms-office.chartstyle+xml";
            else if (type == ContentType::ChartColorStyle)
                typeString = "application/vnd.ms-office.chartcolorstyle+xml";
            else if (type == ContentType::ControlProperties)
                typeString = "application/vnd.ms-excel.controlproperties+xml";
            else if (type == ContentType::CalculationChain)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
            else if (type == ContentType::VBAProject)
                typeString = "application/vnd.ms-office.vbaProject";
            else if (type == ContentType::CoreProperties)
                typeString = "application/vnd.openxmlformats-package.core-properties+xml";
            else if (type == ContentType::ExtendedProperties)
                typeString = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
            else if (type == ContentType::CustomProperties)
                typeString = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
            else if (type == ContentType::Comments)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
            else if (type == ContentType::Table)
                typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
            else if (type == ContentType::VMLDrawing)
                typeString = "application/vnd.openxmlformats-officedocument.vmlDrawing";
            else
                typeString = "Unknown ContentType";
            return typeString;
        }(type);

        auto node = content_types->xmlDocument().first_child().append_child("Override");
        node.append_attribute("PartName").set_value(path.c_str());
        node.append_attribute("ContentType").set_value(typeString.c_str());
    }

    void deleteOverride(const std::string& path)
    {
        content_types->xmlDocument().document_element().remove_child(
            content_types->xmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
    }

    std::vector<XMLNode> getContentItems()
    {
        std::vector<XMLNode> result;
        for (auto item : content_types->xmlDocument().document_element().children()) {
            if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
        }

        return result;
    }
  
#pragma endregion
    static  XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);
 
    static  XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber);
protected:
    FString m_filePath; /**< The path to the original file*/
   // std::filesystem::path path_;

     std::list<_XML_DATA>    m_data{};              /**<  */
     std::deque<std::string> sharedStringCache; /**<  */

     _XML_DATA* shared_strings; 

     _XML_DATA* doc_rels;
     _XML_DATA* workbook_rels;
     _XML_DATA* content_types;
     _XML_DATA* app_properties;
     _XML_DATA* core_properties;
     _XML_DATA* workbook;
 
    Zippy::ZipArchive m_archive; /**< */

	USheet* _curSheet;
};
