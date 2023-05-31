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

#include "Cell.h"
#include "CellValue.h" 
#include "Sheet.h"
#include "Document.h"
#include <regex>

std::regex is_float_pattern(R"(^(?:[+-]?\d+\.\d*(?:[eE][+-]?\d+)?)|(?:[+-]?\d*\.\d+(?:[eE][+-]?\d+)?)$)");
 
FCellReference UCell::GetReference()const
{  
	return { cellNode.attribute("r").value() };
}
 
 
bool UCell::HasFormula() const
{
	if (!*this) return false;
	return cellNode.child("f") != nullptr;
}
  
FString UCell::ToString() const
{
	auto f = cellNode.child("f");
	auto t = cellNode.attribute("t");
	auto tv = t.value();
	auto v = cellNode.child("v");
	if (f)
	{ 
		// ===== If the formula type is 'shared' or 'array', throw an exception.
		auto ft = f.attribute("t");
		if ( ft)
		{	
			auto type = std::string(ft.value());
			if (type == "shared" || type == "array")
			{
				return "";
			} 
		}  

		return FString(f.text().get());
	}
	if (!t && !tv)
	{
		return FCellValue();
	}
	

	if (!tv || strcmp(tv, "n") == 0 || strcmp(tv, "str") == 0 || strcmp(tv, "b") == 0)
	{
		return v.text().get(); 
	}
	else if (strcmp(tv, "s") == 0)
	{
		return sheet->doc_->getString(static_cast<uint32_t>(v.text().as_ullong()));
	} 
	else if (strcmp(tv, "inlineStr") == 0)
	{
		return  cellNode.child("is").child("t").text().get() ;
	}
	return FCellValue();
}
 
 FCellValue UCell::to_value(const USheet* sheet, const XMLNode& cellNode)
{
	// if has formula
	 auto f = cellNode.child("f");
	if (f) return FCellValue(f.text().get());
	auto t = cellNode.attribute("t");
	auto tv = t.value();
	auto v = cellNode.child("v");
	if (!t && !v)
	{
		return FCellValue();
	}

	if (!t || strcmp(tv, "n") == 0)
	{
		std::string numberString =v.text().get();

			std::smatch result;
			bool ret = std::regex_match(numberString, result, is_float_pattern);
			if (ret)
			{
				return FCellValue(std::stof(numberString));
			}
		return FCellValue(std::stoi(numberString));
	} 
	else if (strcmp(tv, "s") == 0)
	{
		return FCellValue(sheet->doc_->getString(static_cast<uint32_t>(v.text().as_ullong())));
	}
	else if (strcmp(tv, "str") == 0)
	{
		return FCellValue{ v.text().get() };
	}
	else if (strcmp(tv, "inlineStr") == 0)
	{
		return FCellValue{ cellNode.child("is").child("t").text().get() };
	}
	else if (strcmp(tv, "b") == 0)
	{
		return FCellValue(v.text().as_bool());
	}

	return FCellValue();
}

  
  
void UCell::SetFormula(FString formula)
{
	auto str = std::string(TCHAR_TO_UTF8(*formula));

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("f")) cellNode.append_child("f");
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== Remove the type and shared index attributes, if they exists.
	cellNode.child("f").remove_attribute("t");
	cellNode.child("f").remove_attribute("si");

	// ===== Set the text of the value node.
	cellNode.child("f").text().set(str.c_str());
	cellNode.child("v").text().set(0);
}
 
bool UCell::HasValue() const
{ 
	if (cellNode.child("v"))
	{
		return true;
	}
	else if (cellNode.child("f"))
	{
		return true;
	}
	auto inlinstr = cellNode.child("is");
	if(inlinstr && inlinstr.child("t"))
	{
		return true;
	}
	return false;
}

void UCell::Clear()
{
	assert(m_cellNode);
	assert(!cellNode.empty());

	
	if (cellNode.child("f"))
	{	
		// ===== Remove the formula node.
		cellNode.remove_child("f"); 
	}
	else
	{
		// ===== Remove the type attribute
		cellNode.remove_attribute("t");

		// ===== Disable space preservation (only relevant for strings).
		cellNode.remove_attribute(" xml:space");

		// ===== Remove the value node.
		cellNode.remove_child("v");
	}
	  
}
 
void UCell::SetBool(bool val)
{
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// ===== If the cell node doesn't have a type child node, create it.
	if (!cellNode.attribute("t")) cellNode.append_attribute("t");

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== Set the type attribute.
	cellNode.attribute("t").set_value("b");

	// ===== Set the text of the value node.
	cellNode.child("v").text().set(val ? 1 : 0);

	// ===== Disable space preservation (only relevant for strings).
	cellNode.child("v").remove_attribute(cellNode.child("v").attribute("xml:space"));
}

void UCell::SetInt(int32 val)
{
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== The type ("t") attribute is not required for number values.
	cellNode.remove_attribute("t");

	// ===== Set the text of the value node.
	cellNode.child("v").text().set(val);

	// ===== Disable space preservation (only relevant for strings).
	cellNode.child("v").remove_attribute(cellNode.child("v").attribute("xml:space"));
}

void UCell::SetString(FString val)
{ 
	
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// ===== If the cell node doesn't have a type child node, create it.
	if (!cellNode.attribute("t")) cellNode.append_attribute("t");

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== Set the type attribute.
	cellNode.attribute("t").set_value("s");

	auto str = std::string(TCHAR_TO_UTF8(*val));
	// ===== Get or create the index in the XLSharedStrings object.
	auto idx = sheet->doc_->getStringIndex(str);
	auto index = (idx>=0 ? idx : sheet->doc_->appendString(str));

	// ===== Set the text of the value node.
	cellNode.child("v").text().set(index);

	////  IMPLEMENTATION FOR EMBEDDED STRINGS:
	// cellNode.attribute("t").set_value("str");
	// cellNode.child("v").text().set(stringValue);
	// 
	// auto s = std::string_view(stringValue);
	// if (s.front() == ' ' || s.back() == ' ') {
	//     if (!cellNode.attribute("xml:space")) cellNode.append_attribute("xml:space");
	//     cellNode.attribute("xml:space").set_value("preserve");
	// }
}





bool UCell::ToBool() const
{
	return cellNode.child("v").text().as_bool();
}

int32 UCell::ToInt()const
{
	return cellNode.child("v").text().as_int();
}
 
#if ENGINE_MAJOR_VERSION>=5
void UCell::Generic_SetFloat(double val)
{
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== The type ("t") attribute is not required for number values.
	cellNode.remove_attribute("t");

	// ===== Set the text of the value node.
	cellNode.child("v").text().set(val);

	// ===== Disable space preservation (only relevant for strings).
	cellNode.child("v").remove_attribute(cellNode.child("v").attribute("xml:space"));
}
void UCell::Generic_ToFloat(double& val)const
{
	val= cellNode.child("v").text().as_double();
}
#else
void UCell::Generic_SetFloat(float val)
{
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// ===== If the cell node doesn't have a value child node, create it.
	if (!cellNode.child("v")) cellNode.append_child("v");

	// ===== The type ("t") attribute is not required for number values.
	cellNode.remove_attribute("t");

	// ===== Set the text of the value node.
	cellNode.child("v").text().set(val);

	// ===== Disable space preservation (only relevant for strings).
	cellNode.child("v").remove_attribute(cellNode.child("v").attribute("xml:space"));
}
void UCell::Generic_ToFloat(float& val)const
{
	val = cellNode.child("v").text().as_float();
}
#endif
 

ECellType UCell::Type()const
{
	// ===== Check that the m_cellNode is valid.
	assert(m_cellNode);              // NOLINT
	assert(!cellNode.empty());    // NOLINT

	// if has formula
	if (cellNode.child("f")) return ECellType::Formula;
	 

	// ===== If neither a Type attribute or a getValue node is present, the cell is empty.
	if (!cellNode.attribute("t") && !cellNode.child("v")) return ECellType::Empty;

	// ===== If a Type attribute is not present, but a value node is, the cell contains a number.
	if ((!cellNode.attribute("t") || (strcmp(cellNode.attribute("t").value(), "n") == 0 && cellNode.child("v") != nullptr))) {
		std::string numberString = cellNode.child("v").text().get();
		if (numberString.find('.') != std::string::npos || numberString.find("E-") != std::string::npos ||
			numberString.find("e-") != std::string::npos)
			return ECellType::Float;
		return ECellType::Integer;
	}

	// ===== If the cell is of type "s", the cell contains a shared string.
	if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "s") == 0)
		return ECellType::String;    // NOLINT

	// ===== If the cell is of type "inlineStr", the cell contains an inline string.
	if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "inlineStr") == 0)
		return ECellType::String;

	// ===== If the cell is of type "str", the cell contains an ordinary string.
	if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "str") == 0)
		return ECellType::String;

	// ===== If the cell is of type "b", the cell contains a boolean.
	if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "b") == 0)
		return ECellType::Boolean;

	// ===== Otherwise, the cell contains an error.
	return ECellType::Error;    // the m_typeAttribute has the ValueAsString "e"
	 
}
void UCell::offset(int32 rowOffset, int32 colOffset)
{
	if (!*this) return;
	auto ref = GetReference();
	FCellReference offsetRef(ref.Row + rowOffset, ref.Col + colOffset);
	auto rownode = UDocument::getRowNode(cellNode.parent().parent(), offsetRef.Row);
	auto cellnode = UDocument::getCellNode(rownode, offsetRef.Col);
	this->cellNode = cellnode;
}

 