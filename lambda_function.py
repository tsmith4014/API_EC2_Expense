import json
import boto3
import requests
from botocore.client import Config
from datetime import datetime, timedelta
from flask import Flask, request, jsonify
from jose import jwt, JWTError
import logging
import os
from dotenv import load_dotenv
from populate_excel import populate_template 

# Load environment variables from .env file
load_dotenv()

# Initialize the Flask application
app = Flask(__name__)

# Cognito configuration from environment variables
USER_POOL_ID = os.getenv('USER_POOL_ID')
APP_CLIENT_ID = os.getenv('CLIENT_ID')
REGION = os.getenv('AWS_REGION')
COGNITO_ISSUER = f"https://cognito-idp.{REGION}.amazonaws.com/{USER_POOL_ID}"
COGNITO_KEYS_URL = f"{COGNITO_ISSUER}/.well-known/jwks.json"

# Function to get Cognito public keys
def get_cognito_public_keys():
    response = requests.get(COGNITO_KEYS_URL)
    response.raise_for_status()
    return response.json()['keys']

COGNITO_PUBLIC_KEYS = get_cognito_public_keys()

def get_cognito_user_id(token):
    try:
        # Decode the JWT token
        header = jwt.get_unverified_header(token)
        key = next((key for key in COGNITO_PUBLIC_KEYS if key['kid'] == header['kid']), None)
        if not key:
            logging.error("Public key not found for kid: %s", header['kid'])
            return None
        user_info = jwt.decode(token, key, algorithms=['RS256'], issuer=COGNITO_ISSUER, audience=APP_CLIENT_ID)
        logging.info("Decoded user info: %s", user_info)
        return user_info['sub']
    except JWTError as e:
        logging.error("JWT Error: %s", e)
        return None

# Setup logging
logging.basicConfig(level=logging.INFO)

@app.route('/')
def home():
    return '<h1>Hello, this server is serving</h1>'

@app.route('/process_expense_report', methods=['POST'])
def process_expense_report():
    try:
        # Get the JWT token from the Authorization header
        auth_header = request.headers.get('Authorization')
        if not auth_header:
            logging.error("Authorization header is missing.")
            return jsonify({'statusCode': 401, 'body': 'Authorization header is missing.'}), 401
        
        token = auth_header.split(' ')[1]
        logging.info("Received token: %s", token)
        cognito_user_id = get_cognito_user_id(token)
        
        if not cognito_user_id:
            logging.error("Invalid token.")
            return jsonify({'statusCode': 401, 'body': 'Invalid token.'}), 401

        # Parse incoming event data
        body = request.json
        logging.info("Received request: %s", body)
        
        period_ending = body['periodEnding']  # YYYY-MM-DD format

        # Extract other form data
        employee_department = body.get('employeeDepartment', 'Default Department')
        school = body.get('school', 'Default School')
        trip_purpose = body.get('tripPurpose', 'Default Purpose')
        travel = body.get('travel', 'No')
        travel_start_date = body.get('travelStartDate', '2022-01-01')
        travel_end_date = body.get('travelEndDate', '2022-01-02')

        # Calculate the start date of the reporting period
        period_ending_date = datetime.strptime(period_ending, '%Y-%m-%d')
        period_start_date = period_ending_date - timedelta(days=6)

        # Query S3 to get files for the user within the date range
        bucket_name = "expensereport-bucket"
        prefix = f"{cognito_user_id}/"
        
        s3_client = boto3.client('s3')
        response = s3_client.list_objects_v2(Bucket=bucket_name, Prefix=prefix)
        if 'Contents' not in response:
            return jsonify({
                'statusCode': 404,
                'body': 'No files found for the user in the specified date range.'
            }), 404
        
        files_data = []
        for obj in response['Contents']:
            key = obj['Key']
            obj_metadata = s3_client.head_object(Bucket=bucket_name, Key=key)['Metadata']
            file_date = datetime.strptime(obj_metadata['date'], '%Y-%m-%d').date()
            if period_start_date.date() <= file_date <= period_ending_date.date():
                files_data.append({
                    'date': file_date,
                    'price': float(obj_metadata['price']),
                    'category': obj_metadata['category']
                })

        # Generate and populate the Excel report
        data = {
            'period_ending': period_ending,
            'files_data': files_data,
            'employee_department': employee_department,
            'school': school,
            'trip_purpose': trip_purpose,
            'travel': travel,
            'travel_start_date': travel_start_date,
            'travel_end_date': travel_end_date
        }
        
        template_path = 'expense_report.xlsx'  # Use the uploaded template path
        output_path = '/tmp/output.xlsx'
        output_path = populate_template(data, template_path, output_path)

        # Upload the file to S3
        s3_file_name = f"{cognito_user_id}/expense_report_{period_ending}.xlsx"
        s3_client.upload_file(output_path, bucket_name, s3_file_name)

        # Generate a presigned URL for the uploaded file
        presigned_url = s3_client.generate_presigned_url('get_object',
                                                        Params={'Bucket': bucket_name, 'Key': s3_file_name},
                                                        ExpiresIn=3600)

        return jsonify({
            'statusCode': 200,
            'body': f'File successfully processed. Download link: {presigned_url}'
        }), 200

    except Exception as e:
        logging.error("Error processing expense report: %s", e)
        return jsonify({
            'statusCode': 500,
            'body': f'Internal server error: {str(e)}'
        }), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
