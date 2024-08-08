import os
import pandas as pd
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from flask import Flask, request, render_template, send_file, redirect, url_for
from datetime import datetime

app = Flask(__name__)

# Read dataset from uploads folder
def read_data(file_path):
    return pd.read_csv(file_path)

# Initialize model with sample data
def initialize_model():
    file_path = "uploads/sales.csv"
    df = read_data(file_path)
    
    # Feature Engineering: Add a Months_Since_Start column
    df['Months_Since_Start'] = df['Year'] * 12 + df['Month']
    
    # Prepare the data
    x = df[['Months_Since_Start']]
    y = df[['Sales']]
    
    # Create and fit the linear regression model
    model = LinearRegression()
    model.fit(x, y)
    return model, df

model, df = initialize_model()

# Function to generate a report for the entire dataset
def generate_report(data, report_type, year=None, month=None):
    if year is not None and month is not None:
        data = data[(data['Year'] < year) | ((data['Year'] == year) & (data['Month'] <= month))]

    predicted_sales = model.predict(data[['Months_Since_Start']])
    data['Predicted Sales'] = predicted_sales
    report = data[['Product', 'Year', 'Month', 'Predicted Sales', 'Sales']]
    
    # Plot the data
    directory = "reports"
    os.makedirs(directory, exist_ok=True)
    
    if report_type == 'bar':
        plt.figure(figsize=(10, 6))  # Smaller size
        plt.bar(data['Months_Since_Start'], data['Sales'], label='Actual Sales', color='blue', edgecolor='black')
        plt.bar(data['Months_Since_Start'], data['Predicted Sales'], label='Predicted Sales', color='red', alpha=0.5, edgecolor='black')
        plt.xlabel('Months Since Start')
        plt.ylabel('Sales')
        plt.title('Bar Graph')
        plt.xticks(data['Months_Since_Start'], ["{}-{:02d}".format(y, m) for y, m in zip(data['Year'], data['Month'])], rotation=90)
        plt.legend()
        plt.grid(True)
        file_name = 'bar_graph.png'
        
    elif report_type == 'dotted':
        plt.figure(figsize=(10, 6))  # Smaller size
        plt.plot(data['Months_Since_Start'], data['Sales'], marker='o', linestyle='-', color='blue', label='Actual Sales')
        plt.plot(data['Months_Since_Start'], data['Predicted Sales'], marker='o', linestyle='--', color='red', label='Predicted Sales')
        plt.xlabel('Months Since Start')
        plt.ylabel('Sales')
        plt.title('Dotted Line Graph')
        plt.xticks(data['Months_Since_Start'], ["{}-{:02d}".format(y, m) for y, m in zip(data['Year'], data['Month'])], rotation=90)
        plt.legend()
        plt.grid(True)
        file_name = 'dotted_graph.png'
        
    else:  # 'combined'
        plt.figure(figsize=(10, 12))  # Smaller size
        
        # Plot Bar Graph
        plt.subplot(2, 1, 1)
        plt.bar(data['Months_Since_Start'], data['Sales'], label='Actual Sales', color='blue', edgecolor='black')
        plt.bar(data['Months_Since_Start'], data['Predicted Sales'], label='Predicted Sales', color='red', alpha=0.5, edgecolor='black')
        plt.xlabel('Months Since Start')
        plt.ylabel('Sales')
        plt.title('Bar Graph')
        plt.xticks(data['Months_Since_Start'], ["{}-{:02d}".format(y, m) for y, m in zip(data['Year'], data['Month'])], rotation=90)
        plt.legend()
        plt.grid(True)
        
        # Plot Dotted Line Graph
        plt.subplot(2, 1, 2)
        plt.plot(data['Months_Since_Start'], data['Sales'], marker='o', linestyle='-', color='blue', label='Actual Sales')
        plt.plot(data['Months_Since_Start'], data['Predicted Sales'], marker='o', linestyle='--', color='red', label='Predicted Sales')
        plt.xlabel('Months Since Start')
        plt.ylabel('Sales')
        plt.title('Dotted Line Graph')
        plt.xticks(data['Months_Since_Start'], ["{}-{:02d}".format(y, m) for y, m in zip(data['Year'], data['Month'])], rotation=90)
        plt.legend()
        plt.grid(True)
        
        plt.tight_layout()  # Adjust layout to avoid overlap
        file_name = 'combined_graph.png'
    
    plot_path = os.path.join(directory, file_name)
    plt.savefig(plot_path, bbox_inches='tight', dpi=300)  # High resolution
    plt.close()  # Close the plot to release resources
    
    return report, plot_path

# Function to write predicted sales to an Excel file
def write_predicted_sales(data):
    predicted_sales = model.predict(data[['Months_Since_Start']])
    data['Predicted Sales'] = predicted_sales

    directory = "reports"
    os.makedirs(directory, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(directory, f"predicted_sales_report_{timestamp}.xlsx")
    
    # Check if the file already exists
    if os.path.exists(file_path):
        # If file exists, append the data
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            data[['Product', 'Year', 'Month', 'Predicted Sales', 'Sales']].to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
    else:
        # If file doesn't exist, create a new Excel file
        data[['Product', 'Year', 'Month', 'Predicted Sales', 'Sales']].to_excel(file_path, index=False)
    
    return file_path

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/predict', methods=['POST'])
def predict():
    year = int(request.form['year'])
    month = int(request.form['month'])
    prediction_data = pd.DataFrame({'Months_Since_Start': [year * 12 + month]})
    predicted_sales = model.predict(prediction_data[['Months_Since_Start']])
    return render_template('result.html', prediction=predicted_sales[0][0])

@app.route('/generate_report', methods=['GET'])
def generate_report_route():
    year = int(request.args.get('year', 2024))  # Default to current year if year is not specified
    month = int(request.args.get('month', 1))  # Default to January if month is not specified
    report_type = request.args.get('type', 'bar')  # Default to bar graph if type is not specified
    report, plot_path = generate_report(df, report_type, year, month)

    # Define the endpoint names for the generated graph
    if report_type == 'bar':
        graph_endpoint = 'generated_bar_graph'
    elif report_type == 'dotted':
        graph_endpoint = 'generated_dotted_graph'
    else:
        graph_endpoint = 'generated_combined_graph'

    return render_template('report.html', report=report.to_html(), plot_path=plot_path, graph_endpoint=graph_endpoint)

@app.route('/generated_bar_graph')
def generated_bar_graph():
    return send_file("reports/bar_graph.png", mimetype='image/png')

@app.route('/generated_dotted_graph')
def generated_dotted_graph():
    return send_file("reports/dotted_graph.png", mimetype='image/png')

@app.route('/generated_combined_graph')
def generated_combined_graph():
    return send_file("reports/combined_graph.png", mimetype='image/png')

@app.route('/write_predicted_sales', methods=['GET'])
def write_predicted_sales_route():
    file_path = write_predicted_sales(df)
    return send_file(file_path, as_attachment=True)

@app.route('/metrics', methods=['GET'])
def metrics():
    metrics_data = {
        "Mean Sales": df['Sales'].mean(),
        "Median Sales": df['Sales'].median(),
        "Standard Deviation": df['Sales'].std()
    }
    metrics_df = pd.DataFrame([metrics_data])
    
    # Generate metrics graph
    plt.figure(figsize=(8, 5))  # Smaller size for better visibility
    plt.bar(metrics_data.keys(), metrics_data.values(), color='skyblue')
    plt.xlabel('Metric')
    plt.ylabel('Value')
    plt.title('Sales Metrics')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    metrics_plot_path = "reports/metrics_graph.png"
    plt.savefig(metrics_plot_path, bbox_inches='tight', dpi=300)  # High resolution
    plt.close()
    
    return render_template('metrics.html', metrics=metrics_df.to_html(), metrics_plot_path=metrics_plot_path)

@app.route('/generated_metrics_graph')
def generated_metrics_graph():
    return send_file("reports/metrics_graph.png", mimetype='image/png')

if __name__ == '__main__':
    app.run(debug=True)
