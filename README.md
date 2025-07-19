# Excel Splitter Web Application

A modern web application that splits master Excel files into multiple separate files based on user-defined rules. Built for Ralph Lauren to automate repetitive Excel processing tasks.

## 🚀 Features

- **Drag & Drop Upload**: Easy file upload with drag & drop support
- **Three Rule Types**: 
  - Single Column Filter
  - AND Logic (two columns)
  - OR Logic (two columns)
- **Smart File Naming**: Auto-generate names or use custom names
- **Individual Downloads**: Download each generated file separately
- **Modern UI**: Beautiful, responsive interface with Bootstrap 5
- **Error Handling**: Comprehensive error handling and user feedback
- **File Safety**: Original files are never modified

## 🛠️ Technology Stack

- **Backend**: Python Flask
- **Frontend**: HTML5, CSS3, JavaScript, Bootstrap 5
- **Excel Processing**: pandas, openpyxl
- **Deployment**: Heroku/Render ready

## 📋 Requirements

- Python 3.8+
- Flask
- pandas
- openpyxl
- gunicorn (for production)

## 🚀 Quick Start

### Local Development

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd excel-splitter
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python app.py
   ```

4. **Open your browser**
   Navigate to `http://localhost:5000`

### Deployment to Heroku

1. **Create Heroku account** (if you don't have one)
   Visit [heroku.com](https://heroku.com)

2. **Install Heroku CLI**
   ```bash
   # Windows
   winget install --id=Heroku.HerokuCLI
   
   # macOS
   brew tap heroku/brew && brew install heroku
   ```

3. **Login to Heroku**
   ```bash
   heroku login
   ```

4. **Create Heroku app**
   ```bash
   heroku create your-app-name
   ```

5. **Deploy to Heroku**
   ```bash
   git add .
   git commit -m "Initial deployment"
   git push heroku main
   ```

6. **Open your app**
   ```bash
   heroku open
   ```

### Deployment to Render

1. **Create Render account** at [render.com](https://render.com)

2. **Connect your GitHub repository**

3. **Create a new Web Service**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`

4. **Deploy automatically**

## 📖 How to Use

### Step 1: Upload Excel File
- Drag and drop your Excel file (.xlsx or .xls) or click to browse
- Maximum file size: 16MB
- The app will analyze your file and show available columns

### Step 2: Define Rules
- **Single Column**: Filter by one column value
- **AND Logic**: Filter by two columns (both conditions must be true)
- **OR Logic**: Filter by two columns (either condition can be true)
- Add custom file names (optional)

### Step 3: Generate Files
- Click "Generate Files" to process your rules
- Download individual Excel files
- Each file contains only the rows that match your rules

## 🎯 Example Use Cases

### Fashion Industry (Ralph Lauren)
- Split product data by season (Spring, Summer, Fall, Winter)
- Filter by gender AND category (Men's Shirts, Women's Dresses)
- Separate by region OR sales channel

### General Business
- Split customer data by location AND age group
- Filter employees by department OR salary range
- Separate orders by status OR date range

## 🔧 Configuration

### Environment Variables
- `SECRET_KEY`: Flask secret key (auto-generated if not set)
- `MAX_FILE_SIZE`: Maximum file size in bytes (default: 16MB)

### Customization
- Modify `app.py` to change business logic
- Update `templates/index.html` for UI changes
- Adjust `requirements.txt` for additional dependencies

## 🛡️ Security Features

- File type validation (.xlsx, .xls only)
- File size limits
- Secure filename handling
- Session-based file management
- Automatic file cleanup

## 📊 Performance

- Handles files up to 16MB
- Processes multiple rules simultaneously
- Efficient pandas operations
- Responsive UI with loading indicators

## 🐛 Troubleshooting

### Common Issues

1. **File upload fails**
   - Check file format (.xlsx or .xls)
   - Ensure file size < 16MB
   - Verify file is not corrupted

2. **No files generated**
   - Check if your rules match any data
   - Verify column names and values
   - Try simpler rules first

3. **Deployment issues**
   - Ensure all files are committed to git
   - Check Heroku/Render logs
   - Verify requirements.txt is correct

### Getting Help

- Check the browser console for JavaScript errors
- Review Flask application logs
- Test with a simple Excel file first

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is developed for Ralph Lauren. All rights reserved.

## 🙏 Acknowledgments

- Built with Flask and Bootstrap
- Excel processing powered by pandas
- Icons from Font Awesome
- Deployed on Heroku/Render

---

**Developed for Ralph Lauren Internship Project** 🏢 