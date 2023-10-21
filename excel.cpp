#include <iostream>
#include <windows.h>
#include <conio.h>
#include <iomanip>
#include <fstream>
#include <string>
#include <vector>
using namespace std;
class Excel
{
    class Cell
    {
        friend class Excel;
        string data;
        Cell *up;
        Cell *down;
        Cell *left;
        Cell *right;

    public:
        Cell(string input = " ", Cell *l = nullptr, Cell *r = nullptr, Cell *u = nullptr, Cell *d = nullptr)
        {
            data = input;
            left = l;
            right = r;
            up = u;
            down = d;
        }
    };
    Cell* current;
	vector<vector<string>> grid;
    int row_size, col_size;
	int c_row, c_col;
};