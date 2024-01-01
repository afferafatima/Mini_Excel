#include <iostream>
#include <vector>
#include <string>
#include <windows.h>
#include <conio.h>
#include <random>
#include <ctime>
#include <fstream>

using namespace std;

struct Cordinate
{
	int row = 0;
	int col = 0;
};

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

			data = to_string(0);
			left = l;
			right = r;
			up = u;
			down = d;
		}
		// temporary function to initialize data  for each cell
		int getRandomNumber()
		{
			srand(time(0)); // Seed the random number generator with the current time
			return rand() % 100;
		}
	};

	Cell *current;
	Cell *head;
	vector<vector<string>> grid;
	int rowSize, colSize;		// row and colomn size at runtime
	int currentRow, currentCol; // current row and current colomn
	/*for excel functions like sum,average*/
	Cell *rangeStart;
	Cordinate range_start{};
	Cordinate range_end{};
	/*to print cell on console*/
	const int cellWidth = 10; // width of each cell at console
	const int cellHeight = 4; // height of each cell at console
	string filename = " ";

	void color(int k)
	{
		HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
		SetConsoleTextAttribute(hConsole, k);
		if (k > 255)
		{
			k = 1;
		}
	}

	void gotoxy(int rpos, int cpos)
	{
		COORD scrn;
		HANDLE hOuput = GetStdHandle(STD_OUTPUT_HANDLE);
		scrn.X = cpos;
		scrn.Y = rpos;
		SetConsoleCursorPosition(hOuput, scrn);
	}

	// print cell
	void printCell(int row, int col, int colour)
	{
		color(colour);
		char c = '*';
		// top horizontal line
		gotoxy(row * cellHeight, col * cellWidth);
		for (int i = 0; i < cellWidth; i++)
		{
			cout << c;
		}

		// Print bottom horizontal line
		gotoxy(row * cellHeight + cellHeight, col * cellWidth);
		for (int i = 0; i <= cellWidth; i++)
		{
			cout << c;
		}

		// Print vertical lines on both sides
		for (int i = 0; i < cellHeight; i++)
		{

			gotoxy(row * cellHeight + i, col * cellWidth);
			cout << c;
		}

		int r = row * cellHeight;
		int ci = (col * cellWidth) + cellWidth;
		for (int i = 0; i <= cellHeight; i++)
		{

			gotoxy(r + i, ci);
			cout << c;
		}

		// Clear the center of the cell
		gotoxy((cellHeight * row) + cellHeight / 2, (col * cellWidth) + 1);
		cout << "     ";
	}

	void printCellData(int row, int col, Cell *d, int colour)
	{
		color(colour);
		gotoxy((cellHeight * row) + cellHeight / 2, (col * cellWidth) + 1);
		cout << d->data;
	}
	// entered value is numeric or not
	bool intCheck(string numeric)
	{
		for (char c : numeric)
		{
			if (!isdigit(c))
			{
				return false;
			}
		}
		return true;
	}

	bool validLength(string data)
	{
		return data.length() < 8;
	}

	bool inputCellData(int row, int col, Cell *d, int colour)
	{
		string data = "";
		color(colour);
		gotoxy((cellHeight * row) + cellHeight / 2, (col * cellWidth) + 1);
		cin >> data;
		if (intCheck(data) && validLength(data))
		{
			d->data = data;
			return true;
		}
		return false;
	}
	void printCol()
	{
		for (int ri = 0, ci = colSize - 1; ri < rowSize; ri++)
		{
			printCell(ri, ci, 7);
		}
	}

	void printRow()
	{
		for (int ci = 0, ri = rowSize - 1; ci < colSize; ci++)
		{
			printCell(ri, ci, 7);
		}
	}

public:
	class iterator
	{
		Cell *t;
		friend class Excel;
		iterator()
		{
			t = nullptr;
		}

		iterator(Cell *data)
		{
			t = data;
		}

		iterator operator++()
		{
			if (t->right != nullptr)
				t = t->right;
			return *this;
		}

		iterator operator++(int)
		{
			if (t->down != nullptr)
				t = t->down;
			return *this;
		}

		iterator operator--()
		{
			if (t->left != nullptr)
				t = t->left;
			return *this;
		}

		iterator operator--(int)
		{
			if (t->right != nullptr)
				t = t->right;
			return *this;
		}

		bool operator==(iterator temp)
		{
			return (t == temp.t);
		}

		bool operator!=(iterator temp)
		{
			return (t != temp.t);
		}

		string &operator*()
		{
			return t->data;
		}
	};

	iterator Get_Head()
	{
		return iterator(head);
	}

	Excel()
	{
		head = nullptr;
		current = nullptr;
		rowSize = colSize = 5;
		currentRow = currentCol = 0;

		head = newRow();
		current = head;
		for (int i = 0; i < rowSize - 1; i++)
		{
			insertRowBelow();
			rowSize--;
			current = current->down;
		}

		current = head;
		printGrid();
		printData();
	}
	Cell *newRow()
	{
		Cell *temp = new Cell(); // cell at the start of the row i.e 0 index
		Cell *curr = temp;
		for (int i = 0; i < colSize - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}
		return curr;
	}
	// insert row below the row containing current cell
	void insertRowBelow()
	{
		Cell *temp = newRow(); // Create a new row of column length
		Cell *temp2 = current; // Store the current cell for navigation

		// Move to the leftmost side of the sheet
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}

		// If it's the last row
		if (temp2->down == nullptr)
		{
			// Link the new row below the last row
			while (temp2 != nullptr)
			{
				temp2->down = temp;
				temp->up = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else // If it's not the last row
		{
			// insert the new row between existing rows
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
		rowSize++; // Increment the row count
	}

	// Create a new row to insert above the row containing the current cell
	void insertRowAbove()
	{
		Cell *temp = newRow(); // Create a new row of cells
		Cell *temp2 = current; // Store the current cell for navigation

		// Move to the leftmost side of the sheet
		while (temp2->left != nullptr)
		{
			temp2 = temp2->left;
		}

		// If the current cell is the leftmost top cell, update the head of the sheet
		if (temp2 == head)
		{
			head = temp;
		}

		// If the new row is inserted above the top row
		if (temp2->up == nullptr)
		{
			// Link the new row above the current row
			while (temp2 != nullptr)
			{
				temp2->up = temp;
				temp->down = temp2;

				temp = temp->right;
				temp2 = temp2->right;
			}
		}
		else // If it's not the top row
		{
			// insert the new row between the existing rows
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

		rowSize++;
		printGrid(); // Update and display the grid
					 // Print_Data(); // print data
	}
	// insert new colomn
	Cell *newCol()
	{
		Cell *temp = new Cell(); // cell at the start of the colomn
		Cell *curr = temp;

		// Create and link new cells for the column
		for (int i = 0; i < rowSize - 1; i++)
		{
			Cell *temp2 = new Cell();
			temp->down = temp2;
			temp2->up = temp;

			temp = temp2;
		}
		return curr;
	}
	// insert colomn to the right of colomn containing currrent cell
	void insertColRight()
	{
		Cell *temp = newCol();
		Cell *temp2 = current;
		// Move to the rightmost side of the sheet
		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}
		// if colomn is inserted right to the right most colomn
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

		colSize++;
	}
	// insert colomn to the right of colomn containing currrent cell
	void insertColLeft()
	{

		Cell *temp = newCol();
		Cell *temp2 = current;

		// Move to the leftmost side of the sheet
		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}
		// If the current cell is the leftmost top cell, update the head of the sheet
		if (temp2 == head)
		{
			head = temp;
		}
		// if colomn is inserted left to left most colomn
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

		colSize++;
		printGrid();
		// Print_Data();
	}
	// insert new cell in a row
	void insertCellByDownShift()
	{
		Cell *temp = current;
		while (current->down != nullptr)
		{
			current = current->down;
		}
		insertRowBelow();
		current = current->down;
		// move data
		while (current != temp)
		{
			current->data = current->up->data;
			current = current->up;
		}
		current->data = " ";
	}
	// insert new cell in a colomn
	void insertCellByRightShift()
	{
		Cell *temp = current;
		while (current->right != nullptr)
		{
			current = current->right;
		}
		insertColRight();
		current = current->right;

		while (current != temp)
		{
			current->data = current->left->data;
			current = current->left;
		}
		current->data = "   ";
	}
	// delete cell data as a row
	void deleteCellByLeftShift()
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
	void deleteCellByUpShift()
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
	// delete current colomn
	void deleteCol()
	{
		if (colSize <= 1) // if only one colomn
			return;
		Cell *temp = current;

		while (temp->up != nullptr)
		{
			temp = temp->up;
		}

		if (temp == head)
		{
			head = temp->right;
		}

		Cell *delete_cell;
		// delete left most colomn
		if (temp->left == nullptr)
		{
			current = current->right;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->right->left = nullptr;

				temp = temp->down;
				delete delete_cell;
			}
		}
		// delete right most colomn
		else if (temp->right == nullptr)
		{
			currentCol--;
			current = current->left;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->left->right = nullptr;
				temp = temp->down;
				delete delete_cell;
			}
		}
		else
		{
			currentCol--;
			current = current->left;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->left->right = temp->right;
				temp->right->left = temp->left;

				temp = temp->down;
				delete delete_cell;
			}
		}
		colSize--;
	}
	// delete current row
	void deleteRow()
	{

		if (rowSize <= 1)
			return;
		Cell *temp = current;

		while (temp->left != nullptr)
		{
			temp = temp->left;
		}

		if (temp == head)
		{
			head = temp->down;
		}

		Cell *delete_cell;
		// delete top most row
		if (temp->up == nullptr)
		{
			current = current->down;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->down->up = nullptr;

				temp = temp->right;
				delete delete_cell;
			}
		}
		// delete bottom row
		else if (temp->down == nullptr)
		{
			currentRow--;
			current = current->up;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->up->down = nullptr;
				temp = temp->right;
				delete delete_cell;
			}
		}
		else
		{
			currentRow--;
			current = current->up;
			while (temp != nullptr)
			{
				delete_cell = temp;
				temp->down->up = temp->up;
				temp->up->down = temp->down;

				temp = temp->right;
				delete delete_cell;
			}
		}
		rowSize--;
	}
	// clear colomn data
	void clearCol()
	{
		Cell *temp = current;
		// goto the top of the colomn to start
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		// clear data
		while (temp != nullptr)
		{
			temp->data = "0";
			temp = temp->down;
		}
	}
	// clear row data
	void clearRow()
	{
		Cell *temp = current;
		// goto the left most side of row to start
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		// clear data
		while (temp != nullptr)
		{
			temp->data = "0";
			temp = temp->right;
		}
	}
	// clear cell data
	void clearCell()
	{
		current->data = "0";
	}
	// print cells of the grid
	void printGrid()
	{

		for (int ri = 0; ri < rowSize; ri++)
		{
			for (int ci = 0; ci < colSize; ci++)
			{
				printCell(ri, ci, 7);
			}
		}
	}
	// print data of each cell
	void printData()
	{
		Cell *temp = head;
		for (int ri = 0; ri < rowSize; ri++)
		{
			Cell *temp2 = temp;
			for (int ci = 0; ci < colSize; ci++)
			{
				printCellData(ri, ci, temp, 7);
				temp = temp->right;
			}

			temp = temp2->down;
		}
	}
	// insert functions in excel call when I pressed
	void insertFunctionsOfExcel()
	{
		char c = _getch();
		if (c == 100 || c == 68) // insert colomn right(d or D )
		{

			insertColRight();

			currentCol++;
			current = current->right;

			printCol();
			printData();
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (c == 97 || c == 65) // insert colomn left(a or A)
		{
			insertColLeft();

			currentCol--;
			current = current->left;

			printCol();
			printData();
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (c == 115 || c == 83) // insert row down(s or S)
		{

			insertRowBelow();

			currentRow++;
			current = current->down;

			printRow();
			printData();
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (c == 119 || c == 87) // insert row up(w or W)
		{

			insertRowAbove();

			currentRow--;
			current = current->up;

			printRow();
			printData();
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (c == 82 || c == 114) // insert cell by right shift(r or R)
		{
			insertCellByRightShift();
			printGrid();
			printData();
		}
		else if (c == 67 || c == 99) // insert cell by down shift(c or C)
		{
			insertCellByDownShift();
			printGrid();
			printData();
		}
		else if (c == 86 || c == 118) // insert cell data (V or v)
		{
			inputCellData(currentRow, currentCol, current, 6);
			printGrid();
			printData();
		}
	}
	// delete functions in excel call when D pressed
	void deleteFunctionsOfExcel()
	{
		char c = _getch();
		if (c == 117 || c == 85) // delete cell by up shift(U  or u)
		{
			deleteCellByUpShift();
			system("cls");
			printGrid();
			printData();
		}
		else if (c == 108 || c == 76) // delete cell by left shift(L  or l)
		{
			deleteCellByLeftShift();
			system("cls");
			printGrid();
			printData();
		}
		else if (c == 82 || c == 114) // delete row(r or R)
		{
			deleteRow();
			system("cls");
			printGrid();
			printData();
		}
		else if (c == 67 || c == 99) // delete col(c or C)
		{
			deleteCol();
			system("cls");
			printGrid();
			printData();
		}
	}
	// clear data of row/colomn/cell
	void clearData()
	{
		char c = _getch();
		if (c == 82 || c == 114) // clear row(r or R)
		{
			clearRow();
			printData();
		}
		else if (c == 67 || c == 99) // clear col(c or C)
		{
			clearCol();
			printData();
		}
		else if (c == 100 || c == 68) // clear cell data(d or D )
		{
			clearCell();
			printData();
		}
	}
	
	// select cells for calculation
	void selectRange()
	{
		rangeStart = current;
		printCell(currentRow, currentCol, 7);
		printCellData(currentRow, currentCol, current, 7);
		int max_col = currentCol, max_row = currentRow, min_col = currentCol, min_row = currentRow;
		while (true)
		{
			char c = _getch();
			if (c == 100) // right(d)
			{
				
				if (current->right != nullptr)
				{
					current = current->right;
					currentCol++;
				}
			}
			else if (c == 97) // left(a)
			{

				if (current->left != nullptr)
				{
					current = current->left;
					currentCol--;
				}
			}
			else if (c == 115) // down(s)
			{
				
				if (current->down != nullptr)
				{
					current = current->down;
					currentRow++;
				}
			}
			else if (c == 119) // up(w)
			{
				if (current->up != nullptr)
				{
					current = current->up;
					currentRow--;
				}
			}
			else if (c == 13) // enter key
				break;
			// update at each loop
			if (currentCol > max_col)
				max_col = currentCol;
			if (currentRow > max_row)
				max_row = currentRow;
			if (currentCol < min_col)
				min_col = currentCol;
			if (currentRow < min_row)
				min_row = currentRow;
		}
		range_end.col = max_col;
		range_end.row = max_row;
		range_start.col = min_col;
		range_start.row = min_row;
	}
	// calculate sum
	int getRangeSum()
	{
		Cell *temp = rangeStart;
		int sum = 0;
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;

		for (int ri = 0; ri <= row_limit; ri++)
		{

			Cell *temp2 = temp;

			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (intCheck(temp->data))
				{

					sum = sum + stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		currentRow = sri;
		currentCol = sci;
		return sum;
	}
	int getRangeAverage()
	{
		Cell *temp = rangeStart;
		int average = 0;
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;

		for (int ri = 0; ri <= row_limit; ri++)
		{

			Cell *temp2 = temp;

			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (intCheck(temp->data))
				{

					average = average + stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		currentRow = sri;
		currentCol = sci;
		return (average / (row_limit * col_limit));
	}
	int getRangeMinimum()
	{
		Cell *temp = rangeStart;
		int minimum = 2000000;
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;

		for (int ri = 0; ri <= row_limit; ri++)
		{

			Cell *temp2 = temp;

			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (stoi(temp->data) < minimum)
				{

					minimum = stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		currentRow = sri;
		currentCol = sci;
		return minimum;
	}
	int getRangeMaximum()
	{
		Cell *temp = rangeStart;
		int maximum = -2000000;
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;

		for (int ri = 0; ri <= row_limit; ri++)
		{

			Cell *temp2 = temp;

			for (int ci = 0; ci <= col_limit; ci++)
			{
				if (stoi(temp->data) > maximum)
				{

					maximum = stoi(temp->data);
				}
				temp = temp->right;
			}
			temp = temp2->down;
		}
		currentRow = sri;
		currentCol = sci;
		return maximum;
	}
	int getRangeCount()
	{
		Cell *temp = rangeStart;
		int count = 0;
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;
		if (row_limit == 0)
			return col_limit + 1;
		else if (col_limit == 0)
			return row_limit + 1;
		else
			return row_limit * col_limit;
	}
	void mathematicalOperations()
	{
		int value = 0;
		char c = getch();
		if (c == 115 || c == 83) // sum(s or S)
		{
			value = getRangeSum();
		}
		else if (c == 97 || c == 65) // average(a or A)
		{
			value = getRangeAverage();
		}
		else if (c == 77 || c == 109) // maximum number(m or M)
		{
			value = getRangeMaximum();
		}
		else if (c == 78 || c == 110) // minimum number(n or N)
		{
			value = getRangeMinimum();
		}
		else if (c == 67 || c == 99) // count(c or c)
		{
			value = getRangeCount();
		}
		else
		{
			return;
		}
		while (true)
		{
			cellMovement();
			char c = _getch();
			if (c == 13) // enter key
				break;
			else if (c == 27) // escape key
				return;
		}
		current->data = to_string(value);
		printCell(currentRow, currentCol, 4);
		printCellData(currentRow, currentCol, current, 6);
	}
	void Copy()
	{
		Cell *temp = rangeStart;
		grid.clear();
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;
		for (int ri = 0; ri <= row_limit; ri++)
		{
			vector<string> clip;
			Cell *temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				clip.push_back(temp->data);
				temp = temp->right;
			}

			grid.push_back(clip);
			temp = temp2->down;
		}

		currentRow = sri;
		currentCol = sci;
	}

	void Cut()
	{
		Cell *temp = rangeStart;
		grid.clear();
		int sri = currentRow;
		int sci = currentCol;

		int row_limit = range_end.row - range_start.row;
		int col_limit = range_end.col - range_start.col;

		for (int ri = 0; ri <= row_limit; ri++)
		{
			vector<string> clip;
			Cell *temp2 = temp;
			for (int ci = 0; ci <= col_limit; ci++)
			{
				clip.push_back(temp->data);
				temp->data = "0";
				temp = temp->right;
				printCellData(currentRow, currentCol, current, 6);
			}

			grid.push_back(clip);
			temp = temp2->down;
		}

		currentRow = sri;
		currentCol = sci;
	}
	void Paste()
	{
		Cell *temp = current;
		for (int ri = 0; ri < grid.size(); ri++)
		{
			Cell *temp2 = current;
			for (int ci = 0; ci < grid[0].size(); ci++)
			{
				current->data = grid[ri][ci];
				printCellData(currentRow, currentCol, current, 6);
				if (current->right == nullptr)
					insertColRight();
				printCol();
				current = current->right;
			}

			if (temp2->down == nullptr)
				insertRowBelow();
			printRow();

			current = temp2->down;
		}

		current = temp;
	}
	void clipBoardFunctions()
	{
		char c = _getch();
		if (c == 99 || c == 67) // copy(c or C)
		{
			Copy();
		}
		else if (c == 120 || c == 88) // cut(x or X)
		{
			Cut();
		}
		else if (c == 118 || c == 86) // paste(v or V)
		{
			Paste();
		}
	}
	void saveFile(string fileName)
	{
		Cell *temp = head;
		ofstream fout(fileName);
		fout << rowSize << endl;
		fout << colSize << endl;
		for (int i = 0; i < rowSize; i++)
		{
			Cell *temp2 = temp;
			getch();
			for (int j = 0; j < colSize; j++)
			{
				if (!intCheck(temp->data))
				{
					fout << " "
						 << ",";
				}
				else
				{
					fout << temp->data << ",";
				}

				temp = temp->right;
			}
			fout << endl;
			temp = temp2->down;
		}
		fout.close();
	}

	void saveFile()
	{
		system("cls");
		string name;
		cout << "Enter name of file : ";
		cin >> name;
		filename=name + ".txt";
		saveFile(filename);

	}
	void loadFile(){
		system("cls");
		string name;
		cout << "Enter name of file : ";
		cin >> name;
		filename=name + ".txt";
		loadFile(filename);
		system("cls");
	}
	void loadFile(string fileName)
	{
		ifstream fin(fileName);
		fin >> rowSize;
		fin >> colSize;

		Cell *temp = head;

		for (int i = 0; i < rowSize; i++)
		{
			Cell *temp2 = temp;
			for (int j = 0; j < colSize; j++)
			{
				string cellData;
				getline(fin, cellData, ',');

				if (!intCheck(cellData))
				{
					temp->data = " ";
				}
				else
				{
					temp->data = cellData;
				}

				temp = temp->right;
			}
			temp = temp2->down;
		}

		fin.close();
	}

	void Keyboard()
	{
		printCell(currentRow, currentCol, 4);
		while (true)
		{
			Sleep(100);
			cellMovement();
			char c = _getch();
			if (c == 73 || c == 105) // insert function call(I or i)
			{
				insertFunctionsOfExcel();
			}
			else if (c == 100 || c == 68) // delete function call(D or d)
			{
				deleteFunctionsOfExcel();
			}
			else if (c == 67 || c == 99) // clear data of row/colomn/cell(c or C)
			{
				clearData();
			}

			else if (c == 13) // enter key to select data range
			{
				selectRange();
			}
			// call clipboard functions
			else if (c == 77 || c == 109) // m or M
				mathematicalOperations();

			else if (c == 88||c==120) // x
				Copy();
			else if (c == 89||c==121) // y
			{
				Cut();
			}
			else if (c == 90||c==122) // z
			{
				Paste();
			}
			else if(c==83||c==115)//ascii value for s or S	
			{
		
				if(filename!=" ")
				{
					saveFile(filename);
				}
				else 
				{
					saveFile();
					printGrid();
					printData();
				}
			}
			else if(c==76||c==108)// l or L
			{
				loadFile();
				printGrid();
				printData();
			}
			else if (c == 27) // escape key to exit
			{
				c = getch();
				if (c == 27)
					break;
			}
		}
	}
	void cellMovement()
	{
		if (GetAsyncKeyState(VK_LEFT) && (currentCol - 1) >= 0) // shift current cell to left
		{
			printCell(currentRow, currentCol, 7); //	unhighlight
			printCellData(currentRow, currentCol, current, 7);
			current = current->left;
			currentCol--;
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (GetAsyncKeyState(VK_RIGHT) && (currentCol + 1) < colSize) // shift current cell to right
		{
			printCell(currentRow, currentCol, 7);
			printCellData(currentRow, currentCol, current, 7);
			current = current->right;
			currentCol++;
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (GetAsyncKeyState(VK_UP) && (currentRow - 1) >= 0) // shift current cell to up
		{

			printCell(currentRow, currentCol, 7);
			printCellData(currentRow, currentCol, current, 7);
			current = current->up;
			currentRow--;
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
		else if (GetAsyncKeyState(VK_DOWN) && (currentRow + 1) < rowSize) // shift current cell to down
		{

			printCell(currentRow, currentCol, 7);
			printCellData(currentRow, currentCol, current, 7);
			current = current->down;
			currentRow++;
			printCell(currentRow, currentCol, 4);
			printCellData(currentRow, currentCol, current, 6);
		}
	}
};
int main()
{
	system("cls");
	Excel data;
	data.Keyboard();
	return 0;
}
