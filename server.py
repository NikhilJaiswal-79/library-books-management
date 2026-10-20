"""
Flask Web Server for Library Book Usage Analysis System.
Exposes REST APIs for Books, Borrowers, Circulation, Analytics, and Visualizations,
while serving the modern single-page web interface.
"""

import os
from datetime import date
from flask import Flask, jsonify, request, send_file, send_from_directory

import database
import operations
import analytics
import seed_data
from config import CHARTS_DIR, DEFAULT_LOAN_DAYS

# Initialize database tables on server startup
database.initialize_database()

# Flask app with static folder configuration
app = Flask(__name__, static_folder="static", static_url_path="/static")

# Ensure charts directory exists
os.makedirs(CHARTS_DIR, exist_ok=True)


# ==================== Helper Functions ====================

def validate_json(required_fields=None):
    """Validate JSON request body and return parsed data or error response."""
    if not request.is_json:
        return None, jsonify({"detail": "Content-Type must be application/json"}), 400
    data = request.get_json(silent=True)
    if data is None:
        return None, jsonify({"detail": "Invalid JSON body"}), 400
    if required_fields:
        missing = [f for f in required_fields if f not in data or data[f] in (None, "")]
        if missing:
            return None, jsonify({"detail": f"Missing required fields: {', '.join(missing)}"}), 400
    return data, None, None


def error_response(message, status_code=400):
    """Standardized error response matching FastAPI's {'detail': ...} format."""
    return jsonify({"detail": message}), status_code


# ==================== Frontend Root ====================

@app.route("/")
def read_root():
    """Serves the main single-page web application."""
    index_file = os.path.join(app.static_folder, "index.html")
    if os.path.exists(index_file):
        return send_file(index_file)
    return "<h1>Library Book Analysis System</h1><p>Frontend static files loading...</p>", 200


# ==================== REST API Endpoints ====================

@app.route("/api/status")
def get_system_status():
    """Returns database backend and system health."""
    conn = database.get_raw_connection()
    conn.close()
    return jsonify({
        "status": "online",
        "database_backend": "MYSQL",
        "today": str(date.today()),
        "default_loan_days": DEFAULT_LOAN_DAYS
    })


@app.route("/api/kpis")
def get_kpis():
    """Returns high-level circulation and stock KPIs."""
    try:
        return jsonify(analytics.get_library_kpis())
    except Exception as e:
        return error_response(str(e), 500)


@app.route("/api/books", methods=["GET"])
def list_books():
    """Retrieves all books."""
    raw = operations.get_all_books()
    books = []
    for r in raw:
        books.append({
            "book_id": r[0],
            "isbn": r[1],
            "title": r[2],
            "author": r[3],
            "category": r[4],
            "copies_available": r[5]
        })
    return jsonify(books)


@app.route("/api/books", methods=["POST"])
def create_book():
    """Registers a new book in the catalogue."""
    data, err_resp, err_code = validate_json(["isbn", "title", "author", "category"])
    if err_resp:
        return err_resp, err_code
    
    success, message = operations.add_book(
        data["isbn"], data["title"], data["author"], data["category"], 
        data.get("copies_available", 1)
    )
    if not success:
        return error_response(message)
    return jsonify({"success": True, "message": message})


@app.route("/api/borrowers", methods=["GET"])
def list_borrowers():
    """Retrieves all registered borrowers."""
    raw = operations.get_all_borrowers()
    borrowers = []
    for r in raw:
        borrowers.append({
            "borrower_id": r[0],
            "borrower_name": r[1],
            "roll_no": r[2]
        })
    return jsonify(borrowers)


@app.route("/api/borrowers", methods=["POST"])
def create_borrower():
    """Registers a new borrower."""
    data, err_resp, err_code = validate_json(["borrower_name", "roll_no"])
    if err_resp:
        return err_resp, err_code
    
    success, message = operations.add_borrower(data["borrower_name"], data["roll_no"])
    if not success:
        return error_response(message)
    return jsonify({"success": True, "message": message})


@app.route("/api/loans/active")
def list_active_loans():
    """Retrieves all active unreturned loans with overdue status."""
    return jsonify(operations.get_active_loans())


@app.route("/api/loans/issue", methods=["POST"])
def issue_book_endpoint():
    """Issues a book to a borrower with validation."""
    data, err_resp, err_code = validate_json(["book_id", "borrower_id"])
    if err_resp:
        return err_resp, err_code
    
    success, message = operations.issue_book(
        data["book_id"], data["borrower_id"], 
        data.get("issue_date"), data.get("loan_days", DEFAULT_LOAN_DAYS)
    )
    if not success:
        return error_response(message)
    return jsonify({"success": True, "message": message})


@app.route("/api/loans/return", methods=["POST"])
def return_book_endpoint():
    """Records the return of an issued book."""
    data, err_resp, err_code = validate_json(["loan_id"])
    if err_resp:
        return err_resp, err_code
    
    success, message = operations.return_book(
        data["loan_id"], data.get("return_date")
    )
    if not success:
        return error_response(message)
    return jsonify({"success": True, "message": message})


# ==================== Analytics Endpoints ====================

@app.route("/api/analytics/popular")
def get_popular_books():
    """Returns top borrowed books."""
    limit = request.args.get("limit", 10, type=int)
    df = analytics.get_most_borrowed_books_df(limit=limit)
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/analytics/overdue")
def get_overdue_loans():
    """Returns all overdue loans."""
    df = analytics.get_overdue_loans_df()
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/analytics/categories")
def get_category_summary():
    """Returns category circulation summary."""
    df = analytics.get_category_summary_df()
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/analytics/trends")
def get_monthly_trends():
    """Returns monthly borrowing trajectory."""
    df = analytics.get_monthly_borrowing_trends_df()
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/analytics/statistics")
def get_loan_statistics():
    """Returns loan duration descriptive statistics (NumPy)."""
    return jsonify(analytics.calculate_loan_statistics())


@app.route("/api/analytics/duration-distribution")
def get_duration_distribution():
    """Returns binned loan duration frequencies (NumPy)."""
    return jsonify(analytics.get_loan_duration_distribution())


@app.route("/api/analytics/top-borrowers")
def get_top_borrowers():
    """Returns most active borrowers."""
    limit = request.args.get("limit", 8, type=int)
    df = analytics.get_top_borrowers_df(limit=limit)
    return jsonify(df.to_dict(orient="records"))


@app.route("/api/analytics/inventory-health")
def get_inventory_health():
    """Returns stock vs checkout allocation."""
    return jsonify(analytics.get_inventory_health())


@app.route("/api/analytics/punctuality")
def get_punctuality():
    """Returns return timeliness breakdown."""
    return jsonify(analytics.get_return_punctuality())


@app.route("/api/seed", methods=["POST"])
def seed_database():
    """Populates or resets sample data."""
    try:
        # Check for optional reset parameter
        reset = True
        if request.is_json:
            data = request.get_json(silent=True)
            if data and "reset" in data:
                reset = data["reset"]
        seed_data.populate_sample_data(force_reset=reset)
        return jsonify({"success": True, "message": "Sample library data successfully populated."})
    except Exception as e:
        return error_response(str(e), 500)


@app.route("/api/charts/generate")
def generate_all_charts():
    """Generates and updates Matplotlib charts."""
    try:
        p1 = analytics.plot_popular_books()
        p2 = analytics.plot_monthly_trends()
        p3 = analytics.plot_category_distribution()
        return jsonify({
            "success": True,
            "charts": {
                "popular_books": "/api/charts/popular_books_chart.png",
                "monthly_trends": "/api/charts/monthly_borrowing_trends.png",
                "category_distribution": "/api/charts/category_distribution_chart.png"
            }
        })
    except Exception as e:
        return error_response(str(e), 500)


@app.route("/api/charts/<chart_name>")
def get_chart_image(chart_name):
    """Serves the generated chart images."""
    file_path = os.path.join(CHARTS_DIR, chart_name)
    if os.path.exists(file_path):
        return send_file(file_path, mimetype="image/png")
    return error_response("Chart not found. Call /api/charts/generate first.", 404)


# ==================== Main Entry Point ====================

if __name__ == "__main__":
    print("\nStarting Library Book Usage Analysis Web Server (Flask)...")
    print("Web Interface URL: http://127.0.0.1:8000")
    # For production use: gunicorn -w 4 -b 127.0.0.1:8000 server:app
    app.run(host="127.0.0.1", port=8000, debug=False, use_reloader=False)