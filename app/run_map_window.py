import os, webview

# Point this to your generated HTML
html_path = os.path.abspath("mymap.html")  # change to your file
webview.create_window("My Map", f"file:///{html_path}")
webview.start()