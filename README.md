# Axe Devtools Accessibility Checker Tool

## Description
This project contains a Java class named `Main` that performs accessibility testing on web pages using Deque's axe-core library in conjunction with Selenium WebDriver and Apache POI for Excel file creation and manipulation.

## Usage
The `Main` class includes a series of methods that:
- Uses Selenium WebDriver to navigate through web pages.
- Sets up the necessary configurations before running the tests.
- Reads configuration properties from a file located at `src/main/resources/config.properties`.
- Performs accessibility checks using the Deque's axe-core library on different user types by visiting specified URLs.
- Captures accessibility violations and populate an Excel report with details including User Type, URL, Violation name, Violation impact, Violation count and HTML target element(s).
- Formats the Excel document to be easily readable.

### Running the Tests
To run the accessibility tests:
1. Ensure that all the dependencies are resolved before trying to run the Maven command.
2. Ensure you have configured the `config.properties` file with the required properties.
3. Execute the `main()` method within the `Main` class using the following Maven command `mvn clean install exec:java`.

## Contributions
Contributions to enhance the functionality, improve reporting, or expand testing capabilities are welcome. Feel free to fork the repository and submit pull requests.

## License
This project is open-source under an MIT License. Refer to the [LICENSE](LICENSE) file for detailed information.