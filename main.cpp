#include <algorithm>
#include <iostream>
#include <tabulate/table.hpp>
#include <xlnt/xlnt.hpp>
using namespace std;
using namespace tabulate;

// Encapsulation concept !
//OOP
// Student Class
class Student {
private:
  string name;
  int age;

public:
  Student(string name, int age) {
    this->name = name;
    this->age = age;
  }
  string getName() const { return name; }
  int getAge() const { return age; }

  void setName(string name) { this->name = name; }
  void setAge(int age) { this->age = age; }
};

//  user class for login system
class User {
private:
  string username;
  string password;
  string role; // Teacher,student,admin
public:
  User(string username, string password, string role) {
    this->username = username;
    this->password = password;
    this->role = role;
  }

  string getUsername() const { return username; }
  string getPassword() const { return password; }
  string getRole() const { return role; }
};

//  Global logged-in user state
User *loggedInUser = nullptr; // nullptr means no user is logged in

// predefined users
vector<User> userDatabase = {
    User("admin", "admin123", "admin"), User("student", "stu123", "student"),
    User("teacher", "tea123", "student") // teacher has same access as student
};

//  printMenu
void printMenu(vector<string> MenuList) {
  Table table;
  table.add_row({"NO", "Menu "});
  for (int i = 0; i < MenuList.size(); i++) {
    table.add_row({to_string(i + 1), MenuList[i]});
  }
  table[0].format().font_style({FontStyle::bold});
  cout << table << endl;
}
//  printStudentTable with pagination

void printStudentTable(vector<Student> studentsList) {
  // Pagintion Settings
  int pageSize = 3;
  int TotalPages = ((int)studentsList.size() + pageSize - 1) / pageSize;
  int currentPage = 1;

  if (studentsList.empty()) {
    cout << "No student data available!! ❌" << endl;
    return;
  }
  while (true) {
    system("clear");

    // Calulate slice for this page
    int start = (currentPage - 1) * pageSize;
    int end = min(start + pageSize, (int)studentsList.size());

    Table table;
    table.add_row({"#", "Name", "Age"});
    table[0].format().font_style({FontStyle::bold});

    for (int i = start; i < end; i++) {
      table.add_row({to_string(i + 1), studentsList[i].getName(),
                     to_string(studentsList[i].getAge())});
    }
    cout << table << endl;

    cout << "Page " << currentPage << " of " << TotalPages << endl;
    cout << "[N]: Next Page | [P]: Previous Page | [Q]: Quit " << endl;
    cout << "Choose your option : ";

    char nav;
    cin >> nav;
    nav = tolower(nav);

    if (nav == 'n' && currentPage < TotalPages) {

      currentPage++;

    } else if (nav == 'p' && currentPage > 1) {

      currentPage--;
    } else if (nav == 'q') {
      break;
    }
  }
}
//  Excel 
void writeVectorToExcel(string filename, vector<Student> students) {
  xlnt::workbook wb;
  auto ws = wb.active_sheet();
  ws.title("Sheet1");

  // table header
  ws.cell("A1").value("Name");
  ws.cell("B1").value("Age");

  // Enter all the data into row
  int row = 2;
  for (auto student : students) {
    ws.cell("A" + to_string(row)).value(student.getName());
    ws.cell("B" + to_string(row)).value(student.getAge());
    row++;
  }
  wb.save(filename);
  cout << "Successfuly saved vector to the excel file" << endl;
}

vector<Student> readStudentsFromExcel(const string &filename) {
  vector<Student> students;
  xlnt::workbook wb;

  try {
    wb.load(filename);
  } catch (...) {
    cout << "Excel file cannot be open for reading!! ❌" << endl;
    return students;
  }

  // open sheet , loop through row and convert to vector
  xlnt::worksheet ws = wb.active_sheet(); // Sheet1
  for (auto row : ws.rows(false)) {
    // skip table header
    if (row[0].to_string() == "Name")
      continue;

    string name;
    int age;
    name = row[0].to_string();
    age = stoi(row[1].to_string());

    Student student(name, age);
    students.push_back(student);
  }
  return students;
}

//  Login function
bool login() {
  system("clear");

  Table banner;
  banner.add_row({"Student Management System"});
  banner[0].format().font_style({FontStyle::bold}).font_color(Color::blue);
  cout << banner << endl;

  cout << "====== Login ======" << endl;
  string username, password;
  cout << "Username: ";
  cin >> username;
  cout << "Password: ";
  cin >> password;

  for (auto &user : userDatabase) {
    if (user.getUsername() == username && user.getPassword() == password) {
      loggedInUser = &user;
      cout << "Login successful!! ✅" << endl;
      return true;
    }
  }
  cout << "Invalid username or password!! ❌" << endl;
  return false;
}
//  Logout function

void logout() {
  cout << "Logging out " << loggedInUser->getUsername() << "... 👋" << endl;
  loggedInUser = nullptr;
}

// SECTION 10: Sort function

void sortStudents(vector<Student> &students) {
  cout << "Sort by: " << endl;
  cout << "1. Name (A-Z)" << endl;
  cout << "2. Name (Z-A)" << endl;
  cout << "3. Age (Low to High)" << endl;
  cout << "4. Age (High to Low)" << endl;
  cout << "Choose your option : ";

  int option;
  cin >> option;

  switch (option) {
  case 1:
    sort(students.begin(), students.end(),
         [](const Student &a, const Student &b) {
           return a.getName() < b.getName();
         });
    cout << "Sorted by Name A-Z ✅" << endl;
    break;
  case 2:
    sort(students.begin(), students.end(),
         [](const Student &a, const Student &b) {
           return a.getName() > b.getName();
         });
    cout << "Sorted by Name Z-A ✅" << endl;
    break;
  case 3:
    sort(students.begin(), students.end(),
         [](const Student &a, const Student &b) {
           return a.getAge() < b.getAge();
         });
    cout << "Sorted by Age Low→High ✅" << endl;
    break;
  case 4:
    sort(students.begin(), students.end(),
         [](const Student &a, const Student &b) {
           return a.getAge() > b.getAge();
         });
    cout << "Sorted by Age High→Low ✅" << endl;
    break;
  default:
    cout << "Invalid choice ❌" << endl;
  }
}

int main() {
  system("clear");
  string filename = "studentdata.xlsx";

  // Login loop
  while (loggedInUser == nullptr) {
    bool success = login();
    if (!success) {
      cout << "Try again? (y/n) : ";
      char choice;
      cin >> choice;
      if (tolower(choice) != 'y')
        return 0;
    }
  }

  vector<Student> studentLists = readStudentsFromExcel(filename);

  // Build menus based on role
  vector<string> programMenu = {
    "Insert new student", // 1 - admin only                           
    "Edit Student info",  // 2 - admin only                                
    "Delete Student",     // 3 - admin only                              
    "Show All Students",  // 4 - everyone                       
    "Search Student",     // 5 - everyone              
    "Sort Students",      // 6 - everyone                   
    "Logout"};

  vector<string> studentMenu = {
      "Show All Students", // 1 - everyone
      "Search Student",    // 2 - everyone
      "Sort Students",     // 3 - everyone
      "Logout"             // 4 - everyone
  };

  bool isAdmin = loggedInUser->getRole() == "admin";

  int option;

  do {
    system("clear");

    // Show role badge 
    cout << "Logged in as: " << loggedInUser->getUsername() << " ["
         << loggedInUser->getRole() << "]" << endl
         << endl;

    if (isAdmin) {
      // admin menu
      printMenu(programMenu);
      cout << "Choose your option : ";
      cin >> option;

      switch (option) {

      //  Enter (admin only)
      case 1: {
        string name;
        int age;
        cout << "Enter Student Name: ";
        cin.ignore();
        getline(cin, name);
        cout << "Enter Student Age : ";
        cin >> age;

        studentLists.push_back(Student(name, age));
        writeVectorToExcel(filename, studentLists);
      } break;

      //  Edit (admin only) 
      case 2: {
        int row;
        cout << "Enter row to edit : ";
        cin >> row;

        if (row < 1 || row > (int)studentLists.size()) {
          cout << "Invalid row number ❌" << endl;
          break;
        }

        string name;
        int age;
        cout << "Enter new Student Name: ";
        cin.ignore();
        getline(cin, name);
        cout << "Enter new Student Age : ";
        cin >> age;

        studentLists[row - 1].setName(name);
        studentLists[row - 1].setAge(age);
        writeVectorToExcel(filename, studentLists);
      } break;

      //  Delete (admin only) 
      
      case 3: {
        string studentName;
        cout << "Enter student name to delete : ";
        cin.ignore();
        getline(cin, studentName);

        auto it = find_if(
            studentLists.begin(), studentLists.end(),
            [&](const Student &obj) { return obj.getName() == studentName; });

        if (it != studentLists.end()) {
          studentLists.erase(it);
          writeVectorToExcel(filename, studentLists);
          cout << "Student deleted successfully ✅" << endl;
        } else {
          cout << "Student not found ❌" << endl;
        }
      } break;

      //  Show all (with pagination) 
      //  FIX: now passes studentLists (was studentList)
      case 4: {
        printStudentTable(studentLists);
      } break;

      //  Search 
      //  FIX: now searches studentLists (was studentList)
      case 5: {
        string searchName;
        cout << "Enter Student Name to search : ";
        cin.ignore();
        getline(cin, searchName);

        bool found = false;
        for (Student &student : studentLists) {
          if (student.getName() == searchName) {
            cout << "Student Found ✅" << endl;
            cout << "Name : " << student.getName() << endl;
            cout << "Age  : " << student.getAge() << endl;
            found = true;
            break;
          }
        }
        if (!found)
          cout << "Student Not Found ❌" << endl;
      } break;

      //  Sort 
      case 6: {
        sortStudents(studentLists);
        // Save sorted order back to Excel
        writeVectorToExcel(filename, studentLists);
      } break;

      //  Logout 
      case 7: {
        logout();
      } break;

      default:
        cout << "Invalid option ❌" << endl;
      }

    } else {
      //  Student and teacher menu
      printMenu(studentMenu);
      cout << "Choose your option : ";
      cin >> option;

      // Remap student menu numbers to shared logic
      switch (option) {

      //  Show all ──────────────────────────
      case 1: {
        printStudentTable(studentLists);
      } break;

      //  Search 
      case 2: {
        string searchName;
        cout << "Enter Student Name to search : ";
        cin.ignore();
        getline(cin, searchName);

        bool found = false;
        for (Student &student : studentLists) {
          if (student.getName() == searchName) {
            cout << "Student Found ✅" << endl;
            cout << "Name : " << student.getName() << endl;
            cout << "Age  : " << student.getAge() << endl;
            found = true;
            break;
          }
        }
        if (!found)
          cout << "Student Not Found ❌" << endl;
      } break;

      //  Sort 
      case 3: {
        sortStudents(studentLists);
      } break;

      //  Logout 
      case 4: {
        logout();
      } break;

      default:
        cout << "Invalid option ❌" << endl;
      }
    }

    // After logout it print re-login instead of exiting
    if (loggedInUser == nullptr) {
      cout << "Login again? (Y/N): ";
      char c;
      cin >> c;
      if (tolower(c) == 'y') {
        while (loggedInUser == nullptr) {
          bool ok = login();
          if (!ok) {
            cout << "Try again? (Y/N): ";
            cin >> c;
            if (tolower(c) != 'y')
              return 0;
          }
        }
        isAdmin = (loggedInUser->getRole() == "admin");
        option = 0; // reset so loop continues
      } else {
        break; // exit program
      }
    }

    if (loggedInUser != nullptr) {
      cout << "\nPress Enter to continue...";
      cin.ignore();
      cin.get();
    }

    // Loop exits when user logs out and chooses not to re-login
  } while (loggedInUser != nullptr);

  return 0;
}