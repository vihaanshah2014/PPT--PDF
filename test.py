import requests

# API endpoint URL
url = "https://b9d9-128-6-37-87.ngrok-free.app/convert"

# Path to the PowerPoint file you want to convert
ppt_file_path = "testing.pptx"

# Open the PowerPoint file in binary mode
with open(ppt_file_path, "rb") as ppt_file:
    # Create a dictionary containing the file
    files = {"file": ppt_file}
    
    try:
        # Send a POST request to the API endpoint with the file
        response = requests.post(url, files=files)
        
        # Check if the request was successful
        if response.status_code == 200:
            # Get the filename from the response headers
            filename = response.headers.get("Content-Disposition").split("filename=")[1].strip('"')
            
            # Save the converted PDF file
            with open(filename, "wb") as pdf_file:
                pdf_file.write(response.content)
            
            print("PowerPoint file converted to PDF successfully.")
        else:
            print("Failed to convert PowerPoint file. Status code:", response.status_code)
    
    except requests.exceptions.RequestException as e:
        print("An error occurred:", e)