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

	Cell *current;
	Cell *head;
	vector<vector<string>> grid;
	int row_size, col_size;	  // row and colomn size at runtime
	int c_row, c_col;		  // current row and current colomn
	const int cellWidth = 10; // width of each cell at console
	const int cellHeight = 4; // height of each cell at console

public:
	Excel()
	{
		head = nullptr;
		current = nullptr;
		row_size = col_size = 5;
		c_row = c_col = 0;

		head = NewRow();
		current = head;
		for (int i = 0; i < row_size - 1; i++)
		{
			InsertRowBelow();
			row_size--;
			current = current->down;
		}

		current = head;
	}
	Cell *NewRow()
	{
		Cell *temp = new Cell(); // cell at the start of the row i.e 0 index
		Cell *curr = temp;
		for (int i = 0; i < col_size - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}
		return curr;
	}
	// Insert row below the row containing current cell
	void InsertRowBelow()
	{
		Cell *temp = NewRow(); // creates a new row of of colomn length
		Cell *temp2 = current; // current cell

		// to go left most side of sheet
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}
		// if last row
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
	// Insert row above the row containing current cell
	void InsertRowAbove()
	{

		Cell *temp = NewRow();
		Cell *temp2 = current;

		// to go left most side of sheet
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}

		if (temp2 == head)
		{
			head = temp;
		}
		// if added new row above the top row
		if (temp2->up == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->up = temp;
				temp->down = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->up = temp2->up;
				temp2->up = temp;
				temp->down = temp2;
				temp->up->down = temp;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		row_size++;
	}

	Cell *NewCol()
	{
		Cell *temp = new Cell(); // cell at the start of the colomn
		Cell *curr = temp;
		for (int i = 0; i < row_size - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->down = temp2;
			temp2->up = temp;

			temp = temp2;
		}

		return curr;
	}
	// insert colomn to the right of colomn containing currrent cell
	void InsertColRight()
	{
		Cell *temp = NewCol();
		Cell *temp2 = current;
		// to go to the top of the colomn
		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}
		// if colomn is added to right to the right most colomn
		if (temp2->right == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->right = temp;
				temp->left = temp2;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->right = temp2->right;
				temp2->right = temp;
				temp->left = temp2;
				temp->right->left = temp;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}

		col_size++;
	}
	// insert colomn to the right of colomn containing currrent cell
	void InsertColLeft()
	{

		Cell *temp = NewCol();
		Cell *temp2 = current;

		// to go to the top
		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}

		if (temp2 == head)
		{
			head = temp;
		}
		// if colomn is added to the left to left most colomn
		if (temp2->left == nullptr)
		{
			while (temp2 != nullptr)
			{
				temp2->left = temp;
				temp->right = temp2;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}
		else
		{
			while (temp2 != nullptr)
			{
				temp->left = temp2->left;
				temp2->left = temp;
				temp->right = temp2;
				temp->left->right = temp;

				temp2 = temp2->down;
				temp = temp->down;
			}
		}

		col_size++;
	}
	// insert new cell in a row
	void InsertCellByDownShift()
	{
		Cell *temp = current;
		while (current->down != nullptr)
		{
			current = current->down;
		}
		InsertRowBelow();
		current = current->down;

		while (current != temp)
		{
			current->data = current->up->data;
			current = current->up;
		}
		current->data = " ";
	}
	// insert new cell in a colomn
	void InsertCellByRightShift()
	{
		Cell *temp = current;
		while (current->right != nullptr)
		{
			current = current->right;
		}
		InsertColRight();
		current = current->right;

		while (current != temp)
		{
			current->data = current->left->data;
			current = current->left;
		}
		current->data = "   ";
	}

	// delete cell data as a row
	void DeleteCellByLeftShift()
	{
		Cell *temp = current;
		temp->data = "    ";
		while (temp->right != nullptr)
		{
			temp->data = temp->right->data;
			temp = temp->right;
		}
		temp->data = "    ";
	}
	// delete cell data as a colom
	void DeleteCellByUpShift()
	{
		Cell *temp = current;
		temp->data = " ";
		while (temp->down != nullptr)
		{
			temp->data = temp->down->data;
			temp = temp->down;
		}
		temp->data = " ";
	}
};