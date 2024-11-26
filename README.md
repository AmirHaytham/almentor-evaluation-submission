# Almentor Evaluation Submission 🤖

<div align="center">

[![Node.js Version][nodejs-image]][nodejs-url]
[![NPM Version][npm-image]][npm-url]
[![License: MIT][license-image]][license-url]

</div>

<p align="center">
  <img src="docs/demo.gif" alt="Demo GIF" width="600px">
</p>

## 📝 Table of Contents
- [About](#about)
- [Getting Started](#getting-started)
- [Installation](#installation)
- [Usage](#usage)
- [Features](#features)
- [Built With](#built-with)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## 🧐 About <a name="about"></a>
This script automates the process of submitting multiple responses to Microsoft Forms. Perfect for teachers/instructors who need to submit evaluations for multiple students efficiently. It handles email extraction, Excel sheet creation with validation rules, and automated form submissions.

## 🏁 Getting Started <a name="getting-started"></a>

### Prerequisites
- Node.js (v14 or higher)
- npm (Node Package Manager)
- Chrome/Chromium browser
- Text editor (VS Code recommended)

## 🔧 Installation <a name="installation"></a>

1. Clone the repository
```bash
git clone https://github.com/AmirHaytham/almentor-evaluation-submission.git
```

2. Navigate to project directory
```bash
cd almentor-evaluation-submission
```

3. Install dependencies
```bash
npm install
```

## 🎈 Usage <a name="usage"></a>

### 1. Prepare Student Data
```javascript
// In automation.js
const text = `
student1@example.com
student2@example.com
// Add more emails...
`;
```

### 2. Generate Excel Sheet
```bash
node automation.js
```

### 3. Fill Excel Data
Open `form_responses.xlsx` and fill:
- Student interaction (0-5)
- Commitment level (0-5)
- Behavior selection
- Feedback notes

### 4. Run Form Submission
```bash
node automation.js
# When prompted, type 'y' to confirm
```

## ✨ Features <a name="features"></a>
- 📧 Email extraction from text
- 📊 Excel sheet generation with validation
- 🤖 Automated form submissions
- 🔄 Retry mechanism for failures
- 📝 Detailed error logging
- 📸 Error screenshots

## 🛠️ Built With <a name="built-with"></a>
- [Puppeteer](https://pptr.dev/) - Web automation
- [ExcelJS](https://github.com/exceljs/exceljs) - Excel handling
- [Node.js](https://nodejs.org/) - Runtime environment

## ⚙️ Configuration <a name="configuration"></a>

### Excel Sheet Structure
| Column | Description | Values |
|--------|-------------|---------|
| Student Email | Email address | Text |
| Teacher Code | Teacher identifier | AH-5352 |
| Interaction | Lecture interaction | 0-5 |
| Commitment | Session commitment | 0-5 |
| Behavior | Student behavior | Dropdown |
| Feedback | Additional notes | Text |

## 💭 Troubleshooting <a name="troubleshooting"></a>

### Common Issues
1. **Excel Busy Error**
```bash
Error: EBUSY: resource busy or locked
# Solution: Close Excel file
```

2. **Navigation Timeout**
```bash
Error: Navigation timeout exceeded
# Solution: Check internet connection
```

## 🤝 Contributing <a name="contributing"></a>
1. Fork the repository
2. Create your feature branch
```bash
git checkout -b feature/AmazingFeature
```
3. Commit changes
```bash
git commit -m 'Add AmazingFeature'
```
4. Push to branch
```bash
git push origin feature/AmazingFeature
```
5. Open a Pull Request

## 📝 License <a name="license"></a>
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.



<!-- MARKDOWN LINKS & IMAGES -->
[nodejs-image]: https://img.shields.io/badge/node-%3E%3D%2014.0.0-brightgreen.svg
[nodejs-url]: https://nodejs.org/
[npm-image]: https://img.shields.io/npm/v/npm.svg
[npm-url]: https://www.npmjs.com/
[license-image]: https://img.shields.io/badge/License-MIT-yellow.svg
[license-url]: https://opensource.org/licenses/MIT
