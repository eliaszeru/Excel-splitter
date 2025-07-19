#!/usr/bin/env python3
"""
Test script to debug file generation issue
"""
import requests
import json
import os

def test_file_generation():
    base_url = "http://localhost:5000"
    
    print("=== Testing File Generation ===")
    
    # Step 1: Upload file
    print("\n1. Uploading file...")
    with open('sample_data.xlsx', 'rb') as f:
        files = {'file': f}
        response = requests.post(f"{base_url}/upload", files=files)
    
    print(f"Upload response status: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        session_id = data.get('session_id')
        print(f"Upload successful: {data.get('total_rows')} rows")
        print(f"Session ID: {session_id}")
    else:
        print(f"Upload failed: {response.text}")
        return
    
    # Step 2: Test session
    print("\n2. Testing session...")
    response = requests.get(f"{base_url}/test-session")
    print(f"Session test status: {response.status_code}")
    if response.status_code == 200:
        session_data = response.json()
        print(f"Session data: {session_data}")
    else:
        print(f"Session test failed: {response.text}")
    
    # Step 3: Process rules
    print("\n3. Processing rules...")
    rules_data = {
        "rules": [
            {
                "rule_type": "single",
                "column1": "Gender",
                "value1": ["Men"],
                "column2": "",
                "value2": [],
                "custom_name": "test_men"
            }
        ],
        "session_id": session_id
    }
    
    response = requests.post(
        f"{base_url}/process",
        headers={'Content-Type': 'application/json'},
        data=json.dumps(rules_data)
    )
    
    print(f"Process response status: {response.status_code}")
    print(f"Process response: {response.text}")
    
    # Step 4: Check if files were created
    print("\n4. Checking uploads folder...")
    if os.path.exists('uploads'):
        files = os.listdir('uploads')
        print(f"Files in uploads folder: {files}")
    else:
        print("Uploads folder does not exist")

if __name__ == "__main__":
    test_file_generation() 