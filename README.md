# Axe Devtools Accessibility Checker

## Description
This project contains a Java class named `AccessibilityCheckerTest` that performs accessibility testing on web pages using Deque's axe-core library in conjunction with Selenium WebDriver and Apache POI for Excel file creation and manipulation.

## Usage
The `AccessibilityCheckerTest` class includes a series of methods that:
- Uses Selenium WebDriver to navigate through web pages.
- Sets up the necessary configurations before running the tests.
- Reads configuration properties from a file located at `src/test/resources/accessibility.properties`.
- Performs accessibility checks using the Deque's axe-core library on different user types by visiting specified URLs.
- Captures accessibility violations and populate an Excel report with details including User Type, URL, Violation name, Violation impact, Violation count and HTML target element(s).
- Formats the Excel document to be easily readable.

### Running the Tests
To run the accessibility tests:
1. Ensure you have configured the `accessibility.properties` file with the required properties like URLs, user types, login credentials, etc.
2. Execute the `startScript()` method within the `AccessibilityCheckerTest` class using the following Maven command `mvn clean test`.

### Pre-requisites
- Ensure that all the dependencies are resolved before trying to run the Maven command.
- Ensure to customize the `accessibility.properties` file and adjust the code according to your project's requirements before running the tests.

## Contributions
Contributions to enhance the functionality, improve reporting, or expand testing capabilities are welcome. Feel free to fork the repository and submit pull requests.

## License
This project is open-source under an MIT License. Refer to the [LICENSE](LICENSE) file for detailed information.