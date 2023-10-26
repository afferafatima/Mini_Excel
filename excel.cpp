#include <iostream>
#include <vector>
#include <string>
#include <windows.h>
#include <conio.h>
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
    int rowSize, colSize;         //row and colomn size at runtime
	int currentRow, currentCol;               //current row and current colomn
	/*for excel functions like sum,average*/
	Cell* RangeStart;
	Cordinate Range_start{};  
	Cordinate Range_end{};
	/*to print cell on console*/
    const int cellWidth = 10;       //width of each cell at console
	const int cellHeight = 4;       //height of each cell at console

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
	
	//print cell
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

		//Clear the center of the cell
		gotoxy((cellHeight * row) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
		cout << "     ";
	}

	void printCellData(int row, int col, Cell* d, int colour)
	{
		color(colour);
		gotoxy((cellHeight * row) + cellHeight / 2, (col * cellWidth) + cellWidth / 2);
		cout << d->data;
	}


	void printCol()
	{
		for (int ri = 0, ci = currentCol + 1; ri < rowSize; ri++)
		{
			printCell(ri, ci, 7);
		}
	}

	void printRow()
	{
		for (int ci = 0, ri = currentRow + 1; ci < colSize; ci++)
		{
			printCell(ri, ci, 7);
		}

	}


public:
class iterator
	{
		Cell* t;
		friend class Excel;
		iterator()
		{
			t = nullptr;
		}

		iterator(Cell* data)
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

		string& operator*()
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
		cout<<cellWidth<<endl;
		cout<<cellHeight<<endl;
		getch();
		printGrid();
	}
    Cell* newRow()
	{
		Cell* temp = new Cell(); //cell at the start of the row i.e 0 index
		Cell* curr = temp;
		for (int i = 0; i < colSize - 1; i++)
		{
			Cell* temp2 = new Cell();
			temp->right = temp2;
			temp2->left = temp;
			temp = temp2;
		}
		return curr;
	}
	//insert row below the row containing current cell 
    void insertRowBelow()
	{
		Cell* temp = newRow(); // Create a new row of column length
		Cell* temp2 = current; // Store the current cell for navigation

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
		Cell* temp = newRow();  // Create a new row of cells
		Cell* temp2 = current;  // Store the current cell for navigation

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
		printGrid();  // Update and display the grid
		// Print_Data(); // print data
	}
	//insert new colomn
	Cell* newCol()
	{
		Cell* temp = new Cell();//cell at the start of the colomn
		Cell* curr = temp;

		// Create and link new cells for the column
		for (int i = 0; i < rowSize - 1; i++)
		{
			Cell* temp2 = new Cell();
			temp->down = temp2;
			temp2->up = temp;

			temp = temp2;
		}

		return curr;
	}
	//insert colomn to the right of colomn containing currrent cell
	void insertColRight()
	{
		Cell* temp = newCol();
		Cell* temp2 = current;
		// Move to the rightmost side of the sheet
		while (temp2->up != nullptr)
		{
			temp2 = temp2->up;
		}
		//if colomn is inserted right to the right most colomn
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
	//insert colomn to the right of colomn containing currrent cell
	void insertColLeft()
	{

		Cell* temp = newCol();
		Cell* temp2 = current;

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
		//if colomn is inserted left to left most colomn
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
	//insert new cell in a row
	void insertCellByDownShift()
	{
		Cell* temp = current;
		while (current->down != nullptr)
		{
			current = current->down;
		}
		insertRowBelow();
		current = current->down;
		//move data
		while (current != temp)
		{
			current->data = current->up->data;
			current = current->up;
		}
		current->data = " ";
	}
	//insert new cell in a colomn
	void insertCellByRightShift()
	{
		Cell* temp = current;
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
	//delete cell data as a row
	void deleteCellByLeftShift()
	{
		Cell* temp = current;
		temp->data = "    ";
		while (temp->right!= nullptr)
		{
			temp->data = temp->right->data;
			temp = temp->right;
		}
		temp->data = "    ";
	}
	//delete cell data as a colom
	void deleteCellByUpShift()
	{
		Cell* temp = current;
		temp->data = " ";
		while (temp->down != nullptr)
		{
			temp->data = temp->down->data;
			temp = temp->down;
		}
		temp->data = " ";
	}
	//delete current colomn
	void deleteCol()
	{
		if (colSize <= 1)//if only one colomn
			return;
		Cell* temp = current;

		while (temp->up != nullptr)
		{
			temp = temp->up;
		}


		if (temp == head)
		{
			head = temp->right;
		}

		Cell* delete_cell;
		//delete left most colomn
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
		//delete right most colomn
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
	//delete current row
	void deleteRow()
	{

		if (rowSize <= 1)
			return;
		Cell* temp = current;

		while (temp->left != nullptr)
		{
			temp = temp->left;
		}


		if (temp == head)
		{
			head = temp->down;
		}

		Cell* delete_cell;
		//delete top most row
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
		//delete bottom row
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
	//clear colomn data
	void clearCol()
	{
		Cell* temp = current;
		//goto the top of the colomn to start
		while (temp->up != nullptr)
		{
			temp = temp->up;
		}
		//clear data
		while (temp != nullptr)
		{
			temp->data="    ";
			temp = temp->down;
		}

	}
	//clear row data
	void clearRow()
	{
		Cell* temp = current;
		//goto the left most side of row to start
		while (temp->left != nullptr)
		{
			temp = temp->left;
		}
		//clear data
		while (temp != nullptr)
		{
			temp->data="    ";
			temp = temp->right;
		}

	}
	
	void printGrid()
	{
		printCell(0, 0, 7);
		getch();
		for (int ri = 0; ri < rowSize; ri++)
		{
			for (int ci = 0; ci < colSize; ci++)
			{
				printCell(ri, ci, 7);
			}
		}
	}

	void Keyboard()
	{
		printCell(currentRow, currentCol, 4);
		Cell* temp = current;
		string input;
		while (true)
		{
			Sleep(100);
			if(GetAsyncKeyState(VK_LEFT))
			{
				if((currentCol-1)>=0)
				{
					current=current->left;
					printCell(currentRow,currentCol,7);
					currentCol--;
					printCell(currentRow,currentCol,4);
				}

			}
			if(GetAsyncKeyState(VK_RIGHT))
			{
				if((currentCol+1)<colSize)
				{
					current=current->right;
					printCell(currentRow,currentCol,7);
					currentCol++;
					printCell(currentRow,currentCol,4);
				}

			}
			if(GetAsyncKeyState(VK_UP))
			{
				if((currentRow-1)>=0)
				{
					current=current->up;
					printCell(currentRow,currentCol,7);
					currentRow--;
					printCell(currentRow,currentCol,4);
				}

			}
			if(GetAsyncKeyState(VK_DOWN))
			{
				if((currentRow+1)<rowSize)
				{
					current=current->down;
					printCell(currentRow,currentCol,7);
					currentRow++;
					printCell(currentRow,currentCol,4);
				}

			}
			
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