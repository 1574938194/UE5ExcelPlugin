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
 
#include "CellIterator.h"
#include "Document.h"
 
bool FCellIterator::operator==(const FCellIterator& right)const
{
    if (dataNode)
    {  
        if (dataNode && right.dataNode)
        {
            return *dataNode == *right.dataNode;
        }
        else if(dataNode || right.dataNode)
        {
            return false;
        }
        else
        {
            return true;
        } 
    }
    return Current == right.Current;
}

FCellIterator& FCellIterator::operator++()
{ 
    FCellReference ref = Current;

    // ===== Determine the cell reference for the next cell.
    if (ref.Col < max.Col)
    {
        ref.Col++;
    }
    else if (ref == max)
    {
        reached = true;
    }
    else
    {
        ref = { ref.Row + 1, min.Col };
    }

    if (dataNode)
    {
        if (reached)
        {
            dataNode = XMLNode();
        }
        else if ((max.Row < ref.Row || (max.Row <= ref.Row && max.Col < ref.Col))
            || ref.Row == Current.Row)
        {
            auto node = dataNode.next_sibling();
            FCellReference xRef(node.attribute("r").value());
            // if cell dosen't exist then create it
            if (!node || (xRef.Row != ref.Row || xRef.Col != ref.Col))
            {
                node = dataNode.parent().insert_child_after("c", dataNode);
                node.append_attribute("r").set_value(ref.to_string().c_str());
            }
            dataNode =  node;
        }
        else if (ref.Row > Current.Row)
        {
            auto rowNode = dataNode.parent().next_sibling();
            if (!rowNode || rowNode.attribute("r").as_ullong() != ref.Row) {
                rowNode = dataNode.parent().parent().insert_child_after("row", dataNode.parent());
                rowNode.append_attribute("r").set_value(ref.Row);
                // getRowNode(*m_dataNode, ref.row());
            }

            dataNode = UDocument::getCellNode(rowNode, ref.Col);
        }
        else
        {
            // error
        }
    }
     
    return *this;
}

FCellIterator FCellIterator::operator++(int)    // NOLINT
{
    auto oldIter(*this);
    ++(*this);
    return oldIter;
}
  
FCellIterator::pointer FCellIterator::get() const
{
    auto ret = NewObject<UCell>();
    ret->cellNode = dataNode;
    return ret;
}
 
void FCellIterator::next_row()
{
    if (Current.Row < max.Row)
    {
        ++Current.Row;
        auto rowNode = dataNode.parent().next_sibling();
        if (!rowNode || rowNode.attribute("r").as_ullong() != Current.Row) {
            rowNode = dataNode.parent().parent().insert_child_after("row", dataNode.parent());
            rowNode.append_attribute("r").set_value(Current.Row);
            // getRowNode(*m_dataNode, ref.row());
        }

        dataNode = UDocument::getCellNode(rowNode, Current.Col);
    }
}