"""
Simple test script to verify the Excel Splitter application
"""
import os
import sys
import subprocess
import time

def test_dependencies():
    """Test if all required dependencies are installed"""
    print("🔍 Testing dependencies...")
    
    required_packages = ['flask', 'pandas', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} - OK")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package} - Missing")
    
    if missing_packages:
        print(f"\n❌ Missing packages: {', '.join(missing_packages)}")
        print("Please install missing packages: pip install -r requirements.txt")
        return False
    
    print("✅ All dependencies are installed!")
    return True

def test_sample_data():
    """Create and test sample data"""
    print("\n📊 Creating sample data...")
    
    try:
        # Import and run the sample data creation
        from create_sample_data import create_sample_data
        filename = create_sample_data()
        
        if os.path.exists(filename):
            print(f"✅ Sample data created: {filename}")
            return True
        else:
            print("❌ Sample data creation failed")
            return False
    except Exception as e:
        print(f"❌ Error creating sample data: {e}")
        return False

def test_flask_app():
    """Test if Flask app can start"""
    print("\n🚀 Testing Flask application...")
    
    try:
        # Import the Flask app
        from app import app
        
        # Test basic routes
        with app.test_client() as client:
            # Test home page
            response = client.get('/')
            if response.status_code == 200:
                print("✅ Home page - OK")
            else:
                print(f"❌ Home page - Error: {response.status_code}")
                return False
            
            # Test upload endpoint (should return error without file)
            response = client.post('/upload')
            if response.status_code == 400:
                print("✅ Upload endpoint - OK (returns error without file as expected)")
            else:
                print(f"❌ Upload endpoint - Unexpected: {response.status_code}")
                return False
        
        print("✅ Flask application test passed!")
        return True
        
    except Exception as e:
        print(f"❌ Flask application test failed: {e}")
        return False

def main():
    """Run all tests"""
    print("🧪 Excel Splitter Application Test Suite")
    print("=" * 50)
    
    # Test 1: Dependencies
    if not test_dependencies():
        return False
    
    # Test 2: Sample data
    if not test_sample_data():
        return False
    
    # Test 3: Flask app
    if not test_flask_app():
        return False
    
    print("\n" + "=" * 50)
    print("🎉 All tests passed! Your application is ready to run.")
    print("\n📋 Next steps:")
    print("1. Run: python app.py")
    print("2. Open: http://localhost:5000")
    print("3. Upload the sample_data.xlsx file")
    print("4. Test the splitting functionality")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 