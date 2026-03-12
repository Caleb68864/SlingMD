# SlingMD Tests

This project contains unit tests for the SlingMD Outlook add-in.

## Running the Tests

### Visual Studio
1. Open the SlingMD solution in Visual Studio
2. Build the solution
3. Open the Test Explorer (Test > Test Explorer)
4. Click "Run All Tests" 

### Command Line
You can run the tests from the command line using the following commands:

```
cd /path/to/SlingMD
dotnet restore
dotnet test SlingMD.Tests/SlingMD.Tests.csproj
```

## Prerequisites

Running `dotnet test` requires **Visual Studio** with the **Office/VSTO development workload** installed. Without it the build fails with `error MSB4019: The imported project "...\Microsoft.VisualStudio.Tools.Office.targets" was not found`. Open the solution in Visual Studio or install that workload via the Visual Studio Installer under *Workloads > Office/SharePoint development*.

## Test Structure

The tests are organized by the components they test:

- `Models/ObsidianSettingsTests.cs`: Tests for settings persistence, loading, corrupt-file fallback
- `Services/AttachmentServiceTests.cs`: Tests for attachment link target and wikilink generation
- `Services/ContactServiceTests.cs`: Tests for contact note creation and managed-block repair
- `Services/EmailProcessorTests.cs`: Tests for metadata building and cache/export-flow guard
- `Services/FileServiceTests.cs`: Tests for file operations and path handling
- `Services/TaskServiceTests.cs`: Tests for task generation, cancel/reset state
- `Services/TemplateServiceTests.cs`: Tests for YAML frontmatter building and escaping
- `Services/ThreadServiceTests.cs`: Tests for thread date parsing and missing-inbox handling

## Adding New Tests

To add new tests:

1. Create a new test class in the appropriate folder
2. Add test methods marked with the `[Fact]` attribute
3. Add the class to the `SlingMD.Tests.csproj` file

## Mocking

The tests use Moq for mocking dependencies. Example:

```csharp
var mockService = new Mock<IService>();
mockService.Setup(s => s.SomeMethod()).Returns("test");
```