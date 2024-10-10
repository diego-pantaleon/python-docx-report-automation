# Python-Docx Report Automation

### Project Brief
This simple automation project arose from the realization that I was spending countless hours on repetitive tasks such as copying data from Excel and manually pasting & formatting it into Word documents. To make this process more efficient and save time, I developed an automation solution using Python (with `pandas` and `python-docx`). It not only allowed me to practice my coding skills but also cut down my manual work by around 83%.

### Key Features
- Automatically generate incentive reports in Word, structured with team-level breakdowns, team lead performance, and employee-specific metrics.
- Use of Python’s `pandas` for data manipulation and `python-docx` to create Word documents programmatically, including tables, headers, and formatting.
- The code is flexible and can handle dynamic datasets with varying numbers of employees per month.

### Notes
- **Data Privacy:** The sample data set included in this repository has been pre-processed to protect sensitive information. All Personally Identifiable Information (PII) has been anonymized, and columns that are not relevant to the project but important to the company have been omitted. These changes do not affect the core functionality or the intended output of the code.
- **Randomized Metrics:** For further protection of company data, the metric values used in this code were randomized using Excel’s `RANDBETWEEN()` function and do not reflect real-world performance.

### Caveats and Additional Information
- **Dependencies:** The project relies on Python libraries such as `pandas` and `python-docx`. You’ll need to install these libraries manually if you don’t already have them. You can install them via:
    ```bash
    pip install pandas python-docx
    ```
- **Document Formatting:** The project assumes A4 page settings and specific margin values for Word output. If you work with different formatting standards, you may need to adjust those parameters.
- **Column Handling:** This project expects pre-formatted data, and the code is tailored for specific column structures. You may need to adapt the column selection or aggregation methods if your dataset differs.
- **File Paths:** Ensure that the path for your custom font (if using one, such as Lato) is properly set within the script. You can use Google Drive integration on platforms like Colab.
- **Error Handling:** Basic error handling is included, but depending on the volume and variability of your data, you may want to implement additional validation or debugging steps.

### Disclaimers
- This code is a general solution for automating reporting but may require customization depending on specific organizational needs. 
- The automation was created for internal use, so certain elements (like specific formatting preferences) may reflect personal or team-specific styles and not general best practices for all organizations.

### Future Improvements
- **Refactoring:** As the code was developed iteratively to solve an immediate problem, there are sections that could benefit from refactoring. Improving the structure, modularity, and efficiency of the code will make it easier to maintain and scale in the future.

### Contact
For any questions or suggestions, feel free to open an issue or contact me directly through this repository.
