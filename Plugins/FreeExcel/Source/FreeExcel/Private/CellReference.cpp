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


#include "CellReference.h"
#include <cmath>
 
FCellReference::FCellReference(FString ref)
{  
	if (!ref.IsEmpty())
	{
		uint64_t letterCount = 0;
		for (auto letter : ref) {
			if (letter >= 65)  // NOLINT
				++letterCount;
			else if (letter <= 57)  // NOLINT
				break;
		}

		auto numberCount = ref.Len() - letterCount;

		Row = FCString::Atoi(*ref.Mid(letterCount, numberCount));

        auto colStr = ref.Mid(0, letterCount);
        int32 result = 0;

        for (int32 i = static_cast<int32>(colStr.Len() - 1), j = 0; i >= 0; --i, ++j)
        {
            result += static_cast<int32>((colStr[static_cast<uint64_t>(i)] - asciiOffset) * std::pow(alphabetSize, j));
        }
         
        Col = result;
	}
	if (Row<1 || Row> max_rows)
    {
		Row = 1;
	}
    if (Col<1 || Col> max_cols)
    {
        Col = 1;
    }
}

std::string FCellReference::to_string() const
{
    std::string _Col;
    if (Col <= alphabetSize)
    {
        // ===== If there is one letter in the Column Name:
        _Col.append(1, char(Col + asciiOffset));
    } 
    else if (Col > alphabetSize && Col <= alphabetSize * (alphabetSize + 1)) {
        // ===== If there are two letters in the Column Name:
        _Col.append(1, char((Col - (alphabetSize + 1)) / alphabetSize + asciiOffset + 1));
        _Col.append(1, char((Col - (alphabetSize + 1)) % alphabetSize + asciiOffset + 1));
    } 
    else {
        // ===== If there is three letters in the Column Name:
        _Col.append(1, char((Col - 703) / (alphabetSize * alphabetSize) + asciiOffset + 1));  
        _Col.append(1, char(((Col - 703) / alphabetSize) % alphabetSize + asciiOffset + 1));
        _Col.append(1, char((Col - 703) % alphabetSize + asciiOffset + 1)); 
    }
 
    return _Col +  std::to_string(Row);
}

