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

#include "Sheet.h"
#include "CellValue.h"
#include "Document.h"


FString USheet::Name() const
{ 
	return FString(doc_->sheetName(id_).c_str());
}
 
void USheet::SetName(FString sheetName)
{ 
    doc_->SetSheetName(id_, sheetName);
}
  
TArray<FCellValue> USheet::GetRowData(int32 row)const
{
	TArray<FCellValue> ret;  
    auto rowNode = UDocument::getRowNode(data_->xmlDocument().first_child().child("sheetData"), row);
    for (auto& it : rowNode)
    {   
        ret.Emplace(UCell::to_value(this,it));
    } 
	return ret;
}
TArray<float> USheet::GetRowFloatData(int32 row ,bool skip0)const
{
	TArray<float> ret;

    auto rowNode = UDocument::getRowNode(data_->xmlDocument().first_child().child("sheetData"), row);
    
    for (auto& it : rowNode)
    {
        if (skip0)
        {
            skip0 = false;
            continue;
        }
        ret.Emplace (it.child("v").text().as_float());
    }
    return ret; 
}
 
UCell* USheet::Cell(FCellReference cellRef) const
{
    auto cellNode = XMLNode();
    auto rowNode = UDocument::getRowNode(data_->xmlDocument().first_child().child("sheetData"), cellRef.Row);

    // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
    if (rowNode.last_child().empty() || FCellReference(rowNode.last_child().attribute("r").value()).Col < cellRef.Col) {
        // if (rowNode.last_child().empty() ||
        // XLCellReference::CoordinatesFromAddress(rowNode.last_child().attribute("r").getValue()).second < columnNumber) {
        rowNode.append_child("c").append_attribute("r").set_value(cellRef.to_string().c_str());
        cellNode = rowNode.last_child();
    }
    // ===== If the requested node is closest to the end, start from the end and search backwards...
    else if (FCellReference(rowNode.last_child().attribute("r").value()).Col - cellRef.Col < cellRef.Col) {
        cellNode = rowNode.last_child();
        while (FCellReference(cellNode.attribute("r").value()).Col > cellRef.Col) cellNode = cellNode.previous_sibling();
        if (FCellReference(cellNode.attribute("r").value()).Col < cellRef.Col) {
            cellNode = rowNode.insert_child_after("c", cellNode);
            cellNode.append_attribute("r").set_value(cellRef.to_string().c_str());
        }
    }
    // ===== Otherwise, start from the beginning
    else {
        cellNode = rowNode.first_child();
        while (FCellReference(cellNode.attribute("r").value()).Col < cellRef.Col) cellNode = cellNode.next_sibling();
        if (FCellReference(cellNode.attribute("r").value()).Col > cellRef.Col) {
            cellNode = rowNode.insert_child_before("c", cellNode);
            cellNode.append_attribute("r").set_value(cellRef.to_string().c_str());
        }
    }
    UCell* ret = NewObject<UCell>();
    ret->sheet = const_cast<USheet*>(this);
    ret->cellNode = cellNode;
    return ret;
}

int32 USheet::GetRowCellCount(int32 row)const
{
    auto sheetDataNode = data_->xmlDocument().first_child().child("sheetData");
    FCellReference ref = { UDocument::getRowNode(sheetDataNode, row).last_child().attribute("r").value() };
    return ref.Col;
}

FCellReference  USheet::LastCell() const noexcept
{
    auto sheetDataNode = data_->xmlDocument().first_child().child("sheetData");
    FCellRange range("A1", sheetDataNode.last_child() ? sheetDataNode.last_child().attribute("r").value() : "A1");
    int32 colNum = 0;
    auto it = range.begin();
    for (int i = range.Min.Row; i < range.Max.Row; ++i)
    {
        colNum = std::max(colNum, it.rowCellCount());
        it.next_row();
    }


    return { range.Max.Row, colNum };
}

void USheet::reset(const _XML_DATA* xmlData)
{
    data_ = const_cast<_XML_DATA*>(xmlData);
    // ===== Read the dimensions of the Sheet and set data members accordingly.
    std::string dimensions = data_->xmlDocument().document_element().child("dimension").attribute("ref").value();
    if (dimensions.find(':') == std::string::npos)
        data_->xmlDocument().document_element().child("dimension").set_value("A1");
    else
        data_->xmlDocument().document_element().child("dimension").set_value(dimensions.substr(dimensions.find(':') + 1).c_str());

    // If Column properties are grouped, divide them into properties for individual Columns.
    if (data_->xmlDocument().first_child().child("cols").type() != pugi::node_null) {
        auto currentNode = data_->xmlDocument().first_child().child("cols").first_child();
        while (currentNode != nullptr) {
            int min = std::stoi(currentNode.attribute("min").value());
            int max = std::stoi(currentNode.attribute("max").value());
            if (min != max) {
                currentNode.attribute("min").set_value(max);
                for (int i = min; i < max; i++) { // NOLINT
                    auto newnode = data_->xmlDocument().first_child().child("cols").insert_child_before("col", currentNode);
                    auto attr = currentNode.first_attribute();
                    while (attr != nullptr) { // NOLINT
                        newnode.append_attribute(attr.name()) = attr.value();
                        attr = attr.next_attribute();
                    }
                    newnode.attribute("min") = i;
                    newnode.attribute("max") = i;
                }
            }
            currentNode = currentNode.next_sibling();
        }
    }
}

void USheet::updateSheetName(const std::string& oldName, const std::string& newName)
{
    // ===== Set up temporary variables
    std::string oldNameTemp = oldName;
    std::string newNameTemp = newName;


    // ===== If the sheet name contains spaces, it should be enclosed in single quotes (')
    if (oldName.find(' ') != std::string::npos) oldNameTemp = "\'" + oldName + "\'";
    if (newName.find(' ') != std::string::npos) newNameTemp = "\'" + newName + "\'";

    // ===== Ensure only sheet names are replaced (references to sheets always ends with a '!')
    oldNameTemp += '!';
    newNameTemp += '!';

    // ===== Iterate through all defined names
    for (auto& row : data_->xmlDocument().document_element().child("sheetData")) {
        for (auto& cell : row.children()) {
            if (!cell.child("f"))
            {
                continue;
            }
            std::string formula = cell.child("f").text().get();

            // ===== Skip if formula contains a '[' and ']' (means that the defined refers to external workbook)
            if (formula.find('[') == std::string::npos && formula.find(']') == std::string::npos) {
                // ===== For all instances of the old sheet name in the formula, replace with the new name.
                while (formula.find(oldNameTemp) != std::string::npos) { // NOLINT
                    formula.replace(formula.find(oldNameTemp), oldNameTemp.length(), newNameTemp);
                }
                cell.child("f").set_value(formula.c_str());
            }
        }
    }
}