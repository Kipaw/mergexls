@app.route('/test_download')
def test_download():
    output_file = io.BytesIO()
    output_file.write(b"Test content for file download")
    output_file.seek(0)
    return send_file(output_file, as_attachment=True, download_name='test_file.txt', mimetype='text/plain')
