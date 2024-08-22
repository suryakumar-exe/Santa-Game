# Santa Game

SantaGame is a C# console application that automates a Secret Santa game. The program reads a list of employees from an Excel file, randomly pairs them, and then saves the results in another Excel file. The project also includes a suite of unit tests using MSTest to ensure that the Secret Santa logic works correctly.

## Features

- **Read Employee Data**: Fetches employee names and email addresses from an Excel file.
- **Secret Santa Matching**: Randomly pairs employees, ensuring that no one is paired with themselves or their previous year's match.
- **Results Export**: Saves the Secret Santa results to an Excel file.
- **Unit Testing**: Ensures the integrity of the Secret Santa algorithm using MSTest.

## Prerequisites

- [.NET 6.0 SDK or later](https://dotnet.microsoft.com/download/dotnet/6.0)
- [EPPlus](https://www.nuget.org/packages/EPPlus) (for handling Excel files)
- MSTest (for unit testing)

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/SantaGame.git
cd SantaGame
dotnet add package EPPlus --version <latest-version>

dotnet build
dotnet run --project SantaGame

dotnet test

SantaGame/
│
├── SantaGame/                      # Main project folder
│   ├── Modal/                      # Contains the Employees class
│   ├── Program.cs                  # Entry point of the application
│   └── Santa.cs                    # Contains the Secret Santa logic
│
├── SantaGame.Test/                 # Test project folder
│   ├── UnitTest1.cs                # MSTest unit tests
│   └── TestHelpers.cs              # (Optional) Helper methods for tests
│
└── README.md                       # Project documentation
