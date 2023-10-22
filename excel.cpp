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
    Cell* head;
	vector<vector<string>> grid;
    int row_size, col_size;         //row and colomn size at runtime
	int c_row, c_col;               //current row and current colomn
    const int cellWidth = 10;       //width of each cell at console
	const int cellHeight = 4;       //height of each cell at console

public:
    Excel()
	{
		head = nullptr;
		current = nullptr;
		row_size = col_size = 5;
		c_row = c_col = 0;

		head = New_row();
		current = head;
		for (int i = 0; i < row_size - 1; i++)
		{
			InsertRowBelow();
			row_size--;
			current = current->down;
		}

		current = head;
		
	}
    Cell* New_row()
	{
		Cell* temp = new Cell(); //cell at the start of the row i.e 0 index
		Cell* curr = temp;
		for (int i = 0; i < col_size - 1; i++)
		{
			Cell* temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}
		return curr;
	}
    void InsertRowBelow()
	{
		Cell* temp = New_row();
		Cell* temp2 = current;

		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}
		if (temp2->down == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->down = temp;
				temp->up = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->down = temp2->down;
				temp2->down = temp;
				temp->up = temp2;
				temp->down->up = temp;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		row_size++;
	}
};