import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import re
import threading
import time

# Global variables
TESSERACT_LOADED = False
PIL_LOADED = False
PYPROJ_LOADED = False
NUMPY_LOADED = False

class App:
    def __init__(self, master):
        self.master = master
        master.title("GNSS Coordinate Distance Calculator")
        
        # Set application icon (if available)
        try:
            if getattr(sys, 'frozen', False):
                # Running as packaged app
                base_dir = os.path.dirname(sys.executable)
                icon_path = os.path.join(base_dir, 'resources', 'icon.ico')
                if os.path.exists(icon_path):
                    master.iconbitmap(icon_path)
        except Exception:
            pass  # Ignore icon errors
            
        # Configure theme colors
        self.bg_color = "#f5f5f5"  # Light gray background
        self.header_color = "#3498db"  # Blue header
        self.btn_color = "#2980b9"  # Darker blue for buttons
        self.btn_text_color = "white"
        self.result_bg = "white"
        self.highlight_color = "#27ae60"  # Green for highlights
        
        # Configure master window
        master.configure(bg=self.bg_color)
        
        # Status bar for loading information
        self.status_frame = tk.Frame(master, bg="#333333", height=25)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Initializing...")
        self.status_bar = tk.Label(self.status_frame, textvariable=self.status_var, 
                                  bd=0, bg="#333333", fg="white", anchor=tk.W, padx=10)
        self.status_bar.pack(side=tk.LEFT, fill=tk.Y)
        
        # Progress indicator
        self.progress_var = tk.StringVar()
        self.progress_var.set("")
        self.progress_label = tk.Label(self.status_frame, textvariable=self.progress_var, 
                                      bg="#333333", fg="#2ecc71", anchor=tk.E, padx=10)
        self.progress_label.pack(side=tk.RIGHT, fill=tk.Y)

        # Header
        self.header_frame = tk.Frame(master, bg=self.header_color, height=60)
        self.header_frame.pack(fill=tk.X)
        
        self.title_label = tk.Label(self.header_frame, text="GNSS Coordinate Distance Calculator", 
                                   font=("Arial", 16, "bold"), bg=self.header_color, fg="white", pady=10)
        self.title_label.pack()

        # Main frame
        self.main_frame = tk.Frame(master, bg=self.bg_color)
        self.main_frame.pack(pady=15, padx=15, fill=tk.BOTH, expand=True)

        self.image1_path = ""
        self.image2_path = ""
        
        # Image selection frame
        self.img_select_frame = tk.LabelFrame(self.main_frame, text="Image Selection", 
                                             bg=self.bg_color, font=("Arial", 10, "bold"))
        self.img_select_frame.pack(fill=tk.X, pady=5)
        
        # Image 1 frame
        self.img1_frame = tk.Frame(self.img_select_frame, bg=self.bg_color)
        self.img1_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.img1_label = tk.Label(self.img1_frame, text="Image 1:", bg=self.bg_color, width=10, anchor=tk.W)
        self.img1_label.pack(side=tk.LEFT, padx=5)
        
        self.img1_path_var = tk.StringVar()
        self.img1_path_var.set("No image selected")
        self.img1_path_label = tk.Label(self.img1_frame, textvariable=self.img1_path_var, 
                                       bg="white", relief=tk.SUNKEN, anchor=tk.W, padx=5)
        self.img1_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.btn1 = tk.Button(self.img1_frame, text="Browse...", command=self.load_image1, 
                             bg=self.btn_color, fg=self.btn_text_color, width=10)
        self.btn1.pack(side=tk.LEFT, padx=5)
        
        # Image 2 frame
        self.img2_frame = tk.Frame(self.img_select_frame, bg=self.bg_color)
        self.img2_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.img2_label = tk.Label(self.img2_frame, text="Image 2:", bg=self.bg_color, width=10, anchor=tk.W)
        self.img2_label.pack(side=tk.LEFT, padx=5)
        
        self.img2_path_var = tk.StringVar()
        self.img2_path_var.set("No image selected")
        self.img2_path_label = tk.Label(self.img2_frame, textvariable=self.img2_path_var, 
                                       bg="white", relief=tk.SUNKEN, anchor=tk.W, padx=5)
        self.img2_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.btn2 = tk.Button(self.img2_frame, text="Browse...", command=self.load_image2, 
                             bg=self.btn_color, fg=self.btn_text_color, width=10)
        self.btn2.pack(side=tk.LEFT, padx=5)
        
        # Compute button
        self.compute_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.compute_frame.pack(fill=tk.X, pady=10)
        
        self.compute_btn = tk.Button(self.compute_frame, text="Compute Distance", command=self.compute_distance, 
                                    bg=self.highlight_color, fg="white", font=("Arial", 11, "bold"), 
                                    height=2, width=20)
        self.compute_btn.pack(pady=5)
        self.compute_btn.config(state=tk.DISABLED)  # Initially disabled

        # Results frame
        self.result_frame = tk.LabelFrame(self.main_frame, text="Results", 
                                         bg=self.bg_color, font=("Arial", 10, "bold"))
        self.result_frame.pack(pady=5, fill=tk.BOTH, expand=True)
        
        # Text box for results
        self.result_text = tk.Text(self.result_frame, height=20, width=80, 
                                  bg=self.result_bg, font=("Consolas", 10))
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar
        self.scrollbar = tk.Scrollbar(self.result_frame, command=self.result_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        self.result_text.config(yscrollcommand=self.scrollbar.set)

        # Configure text tags for formatting
        self.result_text.tag_configure("header", font=("Arial", 11, "bold"), foreground="#2c3e50")
        self.result_text.tag_configure("success", foreground="#27ae60")
        self.result_text.tag_configure("error", foreground="#e74c3c")
        self.result_text.tag_configure("warning", foreground="#f39c12")
        self.result_text.tag_configure("info", foreground="#3498db")
        self.result_text.tag_configure("result", font=("Consolas", 10, "bold"))
        
        # Initial message
        self.result_text.insert(tk.END, "Welcome to GNSS Coordinate Distance Calculator!\n", "header")
        self.result_text.insert(tk.END, "This tool extracts coordinates from images and calculates distances.\n\n", "info")
        self.result_text.insert(tk.END, "Instructions:\n", "header")
        self.result_text.insert(tk.END, "1. Select two images containing GNSS coordinate information\n")
        self.result_text.insert(tk.END, "2. Click 'Compute Distance' to analyze the coordinates\n")
        self.result_text.insert(tk.END, "3. Results will display in this panel\n\n")
        self.result_text.insert(tk.END, "Supported coordinate formats:\n", "header")
        self.result_text.insert(tk.END, "• Labeled format: \"Latitude 49.7220\", \"Longitude 8.7325\"\n")
        self.result_text.insert(tk.END, "• Directional format: \"49.7220° N\", \"8.7325° E\"\n")
        self.result_text.insert(tk.END, "• Base marker section with height values\n\n")
        self.result_text.see(tk.END)

        # Start background loading of modules
        self.load_modules_thread = threading.Thread(target=self.load_modules_in_background)
        self.load_modules_thread.daemon = True
        self.load_modules_thread.start()

    def load_modules_in_background(self):
        """Load heavy modules in background to improve startup time"""
        global TESSERACT_LOADED, PIL_LOADED, PYPROJ_LOADED, NUMPY_LOADED
        
        self.status_var.set("Loading PIL (Image processing)...")
        self.progress_var.set("⏳ 1/4")
        start_time = time.time()
        try:
            from PIL import Image
            PIL_LOADED = True
            self.status_var.set(f"PIL loaded ({time.time()-start_time:.1f}s)")
        except ImportError:
            self.status_var.set("Error loading PIL")
        
        # Load numpy next as it's used by other modules
        start_time = time.time()
        self.status_var.set("Loading NumPy...")
        self.progress_var.set("⏳ 2/4")
        try:
            import numpy as np
            self.np = np
            NUMPY_LOADED = True
            self.status_var.set(f"NumPy loaded ({time.time()-start_time:.1f}s)")
        except ImportError:
            self.status_var.set("Error loading NumPy")
        
        # Load pyproj for coordinate transformations
        start_time = time.time()
        self.status_var.set("Loading PyProj (Coordinate system)...")
        self.progress_var.set("⏳ 3/4")
        try:
            from pyproj import Transformer
            self.Transformer = Transformer
            PYPROJ_LOADED = True
            self.status_var.set(f"PyProj loaded ({time.time()-start_time:.1f}s)")
        except ImportError:
            self.status_var.set("Error loading PyProj")
        
        # Load tesseract last as it's the heaviest
        start_time = time.time()
        self.status_var.set("Loading Tesseract OCR engine...")
        self.progress_var.set("⏳ 4/4")
        try:
            import pytesseract
            self.pytesseract = pytesseract
            
            # Set tesseract path
            if getattr(sys, 'frozen', False):
                # Running as packaged app
                base_dir = os.path.dirname(sys.executable)
                tesseract_path = os.path.join(base_dir, 'Tesseract-OCR', 'tesseract.exe')
                if os.path.exists(tesseract_path):
                    pytesseract.pytesseract.tesseract_cmd = tesseract_path
            else:
                # Running as script
                pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
                
            TESSERACT_LOADED = True
            self.status_var.set(f"Tesseract OCR loaded ({time.time()-start_time:.1f}s)")
        except ImportError:
            self.status_var.set("Error loading Tesseract OCR")
        
        # All modules loaded
        if TESSERACT_LOADED and PIL_LOADED and PYPROJ_LOADED and NUMPY_LOADED:
            self.status_var.set("Ready")
            self.progress_var.set("✅ Ready")
            # Enable compute button if both images are selected
            if self.image1_path and self.image2_path:
                self.compute_btn.config(state=tk.NORMAL)
        else:
            self.status_var.set("Some components failed to load")
            self.progress_var.set("⚠️ Error")

    def load_image1(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.tif")])
        if path:
            self.image1_path = path
            # Update the path display label with just the filename
            filename = os.path.basename(path)
            self.img1_path_var.set(filename)
            self.result_text.insert(tk.END, f"Selected Image 1: {filename}\n", "info")
            self.result_text.see(tk.END)
            # Enable compute button if both images are selected and modules loaded
            if self.image2_path and TESSERACT_LOADED and PYPROJ_LOADED:
                self.compute_btn.config(state=tk.NORMAL)

    def load_image2(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.tif")])
        if path:
            self.image2_path = path
            # Update the path display label with just the filename
            filename = os.path.basename(path)
            self.img2_path_var.set(filename)
            self.result_text.insert(tk.END, f"Selected Image 2: {filename}\n", "info")
            self.result_text.see(tk.END)
            # Enable compute button if both images are selected and modules loaded
            if self.image1_path and TESSERACT_LOADED and PYPROJ_LOADED:
                self.compute_btn.config(state=tk.NORMAL)

    def extract_parameters(self, image_path):
        try:
            # Import required modules
            from PIL import Image
            
            # Open and process image
            img = Image.open(image_path)
            self.result_text.insert(tk.END, f"Processing image: {os.path.basename(image_path)}\n")
            
            # Perform OCR
            try:
                text = self.pytesseract.image_to_string(img, lang='fra')
                self.result_text.insert(tk.END, f"OCR completed\n")
            except Exception as e:
                self.result_text.insert(tk.END, f"OCR ERROR: {str(e)}\n")
                # Try without language specification
                try:
                    self.result_text.insert(tk.END, "Trying OCR without language specification...\n")
                    text = self.pytesseract.image_to_string(img)
                except Exception as e2:
                    self.result_text.insert(tk.END, f"Secondary OCR ERROR: {str(e2)}\n")
                    return None, f"OCR Error: {str(e)} / {str(e2)}"
        except Exception as e:
            self.result_text.insert(tk.END, f"Error processing image: {str(e)}\n")
            return None, f"Image Error: {str(e)}"

        # Extract parameters
        result = {}
        
        # Log the full OCR text for debugging
        self.result_text.insert(tk.END, f"OCR Text Length: {len(text)} characters\n")
        
        # Always look for antenna height first
        # Try multiple patterns for antenna height
        antenna_patterns = [
            r"[Hh]auteur\s+d['']\s*antenne\s*[:\s]*([0-9]+,[0-9]+)\s*m",
            r"[Hh]auteur\s+d['']\s*antenne\s*[:\s]*([0-9]+\.[0-9]+)\s*m",
            r"[Aa]ntenna\s+[Hh]eight\s*[:\s]*([0-9]+,[0-9]+)\s*m",
            r"[Aa]ntenna\s+[Hh]eight\s*[:\s]*([0-9]+\.[0-9]+)\s*m"
        ]
        
        for pattern in antenna_patterns:
            antenna_match = re.search(pattern, text)
            if antenna_match:
                try:
                    antenna_height = float(antenna_match.group(1).replace(",", "."))
                    result['Antenna_Height'] = antenna_height
                    self.result_text.insert(tk.END, f"Found antenna height: {antenna_height}m\n")
                    break
                except ValueError as e:
                    self.result_text.insert(tk.END, f"Error parsing antenna height: {str(e)}\n")
        
        # If no antenna height found yet, try a more general approach
        if 'Antenna_Height' not in result:
            # Look for any small value (0-3m) followed by "m" that could be antenna height
            lines = text.split('\n')
            for line in lines:
                if 'antenne' in line.lower() or 'antenna' in line.lower():
                    height_match = re.search(r"([0-9]+[,\.][0-9]+)\s*m", line)
                    if height_match:
                        try:
                            height_value = float(height_match.group(1).replace(",", "."))
                            if 0 < height_value < 3:  # Reasonable antenna height range
                                result['Antenna_Height'] = height_value
                                self.result_text.insert(tk.END, f"Found antenna height from context: {height_value}m\n")
                                break
                        except ValueError:
                            pass

        # Try to extract using explicit labels (original method)
        # Latitude
        lat_match = re.search(r"Latitude\s+([0-9.,°]+)", text)
        if lat_match:
            try:
                result['Latitude'] = float(lat_match.group(1).replace(",", ".").replace("°", ""))
                self.result_text.insert(tk.END, f"Found Latitude with explicit label: {result['Latitude']}\n")
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Latitude: {str(e)}\n")
        
        # Longitude
        lon_match = re.search(r"Longitude\s+([0-9.,°]+)", text)
        if lon_match:
            try:
                result['Longitude'] = float(lon_match.group(1).replace(",", ".").replace("°", ""))
                self.result_text.insert(tk.END, f"Found Longitude with explicit label: {result['Longitude']}\n")
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Longitude: {str(e)}\n")
        
        # Ellipsoidal Height
        ellip_match = re.search(r"Hauteur ellipsoïdale\s+([\d.,]+)\s*m", text)
        if ellip_match:
            try:
                result['Ellipsoidal_Height'] = float(ellip_match.group(1).replace(",", "."))
                self.result_text.insert(tk.END, f"Found Ellipsoidal Height with explicit label: {result['Ellipsoidal_Height']}\n")
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Ellipsoidal Height: {str(e)}\n")

        # If we didn't find coordinates with explicit labels, try alternative formats
        if 'Latitude' not in result or 'Longitude' not in result:
            self.result_text.insert(tk.END, "Trying alternative coordinate formats...\n")
            
            # Look for coordinate pairs like "49,72201469° N" or "8,73250085° E"
            # First try to find N/S coordinates (latitude)
            lat_alt_match = re.search(r"([0-9][0-9],[0-9]+)°\s*[Nn]", text)
            if lat_alt_match:
                try:
                    result['Latitude'] = float(lat_alt_match.group(1).replace(",", "."))
                    self.result_text.insert(tk.END, f"Found N latitude: {result['Latitude']}\n")
                except ValueError as e:
                    self.result_text.insert(tk.END, f"Error parsing N latitude: {str(e)}\n")
            else:
                # Try South latitude
                lat_alt_match = re.search(r"([0-9][0-9],[0-9]+)°\s*[Ss]", text)
                if lat_alt_match:
                    try:
                        result['Latitude'] = -float(lat_alt_match.group(1).replace(",", "."))
                        self.result_text.insert(tk.END, f"Found S latitude: {result['Latitude']}\n")
                    except ValueError as e:
                        self.result_text.insert(tk.END, f"Error parsing S latitude: {str(e)}\n")
            
            # Then try to find E/W coordinates (longitude)
            lon_alt_match = re.search(r"([0-9][0-9]?,[0-9]+)°\s*[Ee]", text)
            if lon_alt_match:
                try:
                    result['Longitude'] = float(lon_alt_match.group(1).replace(",", "."))
                    self.result_text.insert(tk.END, f"Found E longitude: {result['Longitude']}\n")
                except ValueError as e:
                    self.result_text.insert(tk.END, f"Error parsing E longitude: {str(e)}\n")
            else:
                # Try West longitude
                lon_alt_match = re.search(r"([0-9][0-9]?,[0-9]+)°\s*[Ww]", text)
                if lon_alt_match:
                    try:
                        result['Longitude'] = -float(lon_alt_match.group(1).replace(",", "."))
                        self.result_text.insert(tk.END, f"Found W longitude: {result['Longitude']}\n")
                    except ValueError as e:
                        self.result_text.insert(tk.END, f"Error parsing W longitude: {str(e)}\n")
                        
            # Try alternative pattern where coordinates are presented in a list format
            # Look for patterns like "8,73250085° E" and "49,72201469° N" on separate lines
            coord_pattern = re.findall(r"([0-9][0-9]?,[0-9]+)°\s*([EeWwNnSs])", text)
            if coord_pattern and len(coord_pattern) >= 2:
                self.result_text.insert(tk.END, f"Found coordinate pattern list: {coord_pattern}\n")
                for coord in coord_pattern:
                    value = float(coord[0].replace(",", "."))
                    direction = coord[1].upper()
                    
                    if direction in ['N', 'S']:
                        if direction == 'S':
                            value = -value
                        if 'Latitude' not in result:
                            result['Latitude'] = value
                            self.result_text.insert(tk.END, f"Set Latitude from list: {value}\n")
                    elif direction in ['E', 'W']:
                        if direction == 'W':
                            value = -value
                        if 'Longitude' not in result:
                            result['Longitude'] = value
                            self.result_text.insert(tk.END, f"Set Longitude from list: {value}\n")
        
        # Check for antenna height specifically to avoid confusing it with ellipsoidal height
        if 'Antenna_Height' not in result:
            antenna_height_match = re.search(r"[Hh]auteur\s+d['']\s*antenne\s*[:\s]*([0-9]+,[0-9]+)\s*m", text)
            if antenna_height_match:
                try:
                    antenna_height = float(antenna_height_match.group(1).replace(",", "."))
                    result['Antenna_Height'] = antenna_height
                    self.result_text.insert(tk.END, f"Found antenna height: {antenna_height}m\n")
                except ValueError as e:
                    self.result_text.insert(tk.END, f"Error parsing antenna height: {str(e)}\n")

        # Look for height after coordinates in marker/base sections
        if 'Ellipsoidal_Height' not in result and 'Latitude' in result and 'Longitude' in result:
            # Try to find a height value that appears right after coordinates
            # This pattern looks for a line with just a number and "m" after coordinates
            lines = text.split('\n')
            coord_found = False
            
            for i, line in enumerate(lines):
                # Mark when we've found a coordinate line
                if ('N' in line and str(result['Latitude']) in line.replace(',', '.')) or \
                   ('E' in line and str(result['Longitude']) in line.replace(',', '.')):
                    coord_found = True
                    continue
                
                # Look for height in the line after coordinates
                if coord_found:
                    height_match = re.search(r"^\s*([0-9]+,[0-9]+)\s*m", line)
                    if height_match:
                        try:
                            height_value = float(height_match.group(1).replace(",", "."))
                            if 100 <= height_value < 10000:  # Reasonable ellipsoidal height range
                                result['Ellipsoidal_Height'] = height_value
                                self.result_text.insert(tk.END, f"Found ellipsoidal height after coordinates: {height_value}m\n")
                                break
                        except ValueError:
                            pass
                    coord_found = False  # Reset for next coordinate section
        
        # If still no height, look for a pattern like "415,943 m"
        if 'Ellipsoidal_Height' not in result:
            # Look for "Marqueur de base" section which often contains coordinates and height
            base_marker_section = False
            height_after_coords = None
            
            lines = text.split('\n')
            for i, line in enumerate(lines):
                if "Marqueur de base" in line:
                    base_marker_section = True
                    continue
                
                if base_marker_section and i+2 < len(lines):
                    # Try to extract three values (typically lat, lon, height)
                    # The third value is often the ellipsoidal height
                    try:
                        # Skip to last line in this section which is often height
                        for j in range(1, 4):
                            if i+j < len(lines) and re.match(r"^\s*([0-9]+,[0-9]+)\s*m?\s*$", lines[i+j].strip()):
                                height_value = float(lines[i+j].strip().replace("m", "").replace(",", ".").strip())
                                if 100 <= height_value < 10000:  # Reasonable ellipsoidal height range
                                    height_after_coords = height_value
                                    break
                    except ValueError:
                        pass
                    break  # Stop after processing the base marker section
            
            if height_after_coords:
                result['Ellipsoidal_Height'] = height_after_coords
                self.result_text.insert(tk.END, f"Found ellipsoidal height in base marker section: {height_after_coords}m\n")
            else:
                # General search for height values if still not found
                height_values = re.findall(r"([0-9]+,[0-9]+)\s*m", text)
                for height_str in height_values:
                    try:
                        height_value = float(height_str.replace(",", "."))
                        # Filter out small values (likely antenna heights)
                        # and use large values (likely ellipsoidal heights)
                        if 100 <= height_value < 10000 and 'Ellipsoidal_Height' not in result:
                            result['Ellipsoidal_Height'] = height_value
                            self.result_text.insert(tk.END, f"Found ellipsoidal height from general search: {height_value}m\n")
                            break
                    except ValueError:
                        continue

        return result, text

        return result, text

    def determine_utm_epsg(self, lat, lon):
        zone_number = int((lon + 180) / 6) + 1
        if lat >= 0:
            epsg = 32600 + zone_number
        else:
            epsg = 32700 + zone_number
        return epsg

    def latlon_to_utm(self, lat, lon, transformer):
        x, y = transformer.transform(lon, lat)
        return x, y

    def compute_distance(self):
        # Disable the button during computation
        self.compute_btn.config(state=tk.DISABLED)
        self.result_text.delete(1.0, tk.END)
        
        # Show tesseract info
        if TESSERACT_LOADED:
            self.result_text.insert(tk.END, "GNSS Coordinate Analysis\n", "header")
            self.result_text.insert(tk.END, f"Using Tesseract OCR: {os.path.basename(self.pytesseract.pytesseract.tesseract_cmd)}\n\n", "info")
        else:
            self.result_text.insert(tk.END, "ERROR: Tesseract OCR not loaded\n", "error")
            self.compute_btn.config(state=tk.NORMAL)
            return

        if not self.image1_path or not self.image2_path:
            messagebox.showerror("Error", "Please select both images first.")
            self.compute_btn.config(state=tk.NORMAL)
            return

        # Use a thread for processing to keep UI responsive
        threading.Thread(target=self._compute_distance_thread).start()
        
    def _compute_distance_thread(self):
        # Update status during processing
        self.progress_var.set("⏳ Processing...")
        
        # Extract parameters
        self.status_var.set("Processing Image 1...")
        data1, raw1 = self.extract_parameters(self.image1_path)
        
        self.status_var.set("Processing Image 2...")
        data2, raw2 = self.extract_parameters(self.image2_path)

        if not data1 or not data2:
            self.result_text.insert(tk.END, "\nOCR extraction failed. Check images and text layout.\n", "error")
            # Only show raw OCR text when extraction completely fails
            if raw1:
                self.result_text.insert(tk.END, f"\nRaw OCR Text 1:\n{raw1}\n\n")
            if raw2:
                self.result_text.insert(tk.END, f"Raw OCR Text 2:\n{raw2}\n")
            self.compute_btn.config(state=tk.NORMAL)
            self.status_var.set("Ready")
            self.progress_var.set("❌ Failed")
            return

        # Check if all required fields are present
        required_fields = ['Latitude', 'Longitude']
        missing_fields = False
        for field in required_fields:
            if field not in data1:
                self.result_text.insert(tk.END, f"Missing field in Image 1: {field}\n", "error")
                missing_fields = True
            if field not in data2:
                self.result_text.insert(tk.END, f"Missing field in Image 2: {field}\n", "error")
                missing_fields = True
        
        # Special handling for ellipsoidal height - use a default if not found
        if 'Ellipsoidal_Height' not in data1:
            self.result_text.insert(tk.END, "Warning: Ellipsoidal Height not found in Image 1. Using default value of 0.\n", "warning")
            data1['Ellipsoidal_Height'] = 0
        
        if 'Ellipsoidal_Height' not in data2:
            self.result_text.insert(tk.END, "Warning: Ellipsoidal Height not found in Image 2. Using default value of 0.\n", "warning")
            data2['Ellipsoidal_Height'] = 0
        
        if missing_fields:
            self.result_text.insert(tk.END, f"\nImage 1 data: {data1}\n")
            self.result_text.insert(tk.END, f"Image 2 data: {data2}\n")
            
            # Only show raw OCR text when critical fields are missing
            self.result_text.insert(tk.END, f"\nRaw OCR Text 1:\n{raw1}\n\n")
            self.result_text.insert(tk.END, f"Raw OCR Text 2:\n{raw2}\n")
            
            self.compute_btn.config(state=tk.NORMAL)
            self.status_var.set("Ready")
            self.progress_var.set("❌ Failed")
            return

        try:
            self.status_var.set("Calculating coordinates...")
            
            # Determine UTM zone once from Image 1
            epsg_code = self.determine_utm_epsg(data1['Latitude'], data1['Longitude'])
            transformer = self.Transformer.from_crs("EPSG:4326", f"EPSG:{epsg_code}", always_xy=True)

            # Convert
            xy1 = self.latlon_to_utm(data1['Latitude'], data1['Longitude'], transformer)
            xy2 = self.latlon_to_utm(data2['Latitude'], data2['Longitude'], transformer)

            # Distance
            xy_distance = self.np.linalg.norm(self.np.array(xy1) - self.np.array(xy2))
            height_diff = abs(data1['Ellipsoidal_Height'] - data2['Ellipsoidal_Height'])

            # Format the data dictionaries for display, including antenna height if available
            display_data1 = {
                "Latitude": data1['Latitude'],
                "Longitude": data1['Longitude'],
                "Ellipsoidal_Height": data1['Ellipsoidal_Height']
            }
            if 'Antenna_Height' in data1:
                display_data1["Antenna_Height"] = data1['Antenna_Height']
                
            display_data2 = {
                "Latitude": data2['Latitude'],
                "Longitude": data2['Longitude'],
                "Ellipsoidal_Height": data2['Ellipsoidal_Height']
            }
            if 'Antenna_Height' in data2:
                display_data2["Antenna_Height"] = data2['Antenna_Height']

            # Clear previous results and show new results with improved formatting
            self.result_text.delete(1.0, tk.END)
            
            # Add results title with underline
            self.result_text.insert(tk.END, "GNSS COORDINATE ANALYSIS RESULTS\n", "header")
            self.result_text.insert(tk.END, "===============================\n\n")
            
            # Side-by-side comparison using fixed width formatting
            name1 = os.path.basename(self.image1_path)
            name2 = os.path.basename(self.image2_path)
            
            # Section headers with clear separation
            self.result_text.insert(tk.END, "IMAGE 1\n", "header")
            self.result_text.insert(tk.END, "-------\n")
            self.result_text.insert(tk.END, f"File: {name1}\n\n", "info")
            
            # WGS84 Coordinates section
            self.result_text.insert(tk.END, "WGS84 COORDINATES:\n", "header")
            self.result_text.insert(tk.END, f"  Latitude:  {data1['Latitude']:.8f}°\n")
            self.result_text.insert(tk.END, f"  Longitude: {data1['Longitude']:.8f}°\n")
            self.result_text.insert(tk.END, f"  Ellip. Ht: {data1['Ellipsoidal_Height']:.3f} m\n")
            if 'Antenna_Height' in display_data1:
                self.result_text.insert(tk.END, f"  Antenna Ht: {display_data1['Antenna_Height']:.3f} m\n")
            
            # UTM Coordinates section
            self.result_text.insert(tk.END, f"\nUTM COORDINATES (EPSG:{epsg_code}):\n", "header")
            self.result_text.insert(tk.END, f"  Easting:  {xy1[0]:.3f} m\n")
            self.result_text.insert(tk.END, f"  Northing: {xy1[1]:.3f} m\n\n")
            
            # Divider between images
            self.result_text.insert(tk.END, "-" * 40 + "\n\n")
            
            # IMAGE 2 section
            self.result_text.insert(tk.END, "IMAGE 2\n", "header")
            self.result_text.insert(tk.END, "-------\n")
            self.result_text.insert(tk.END, f"File: {name2}\n\n", "info")
            
            # WGS84 Coordinates section
            self.result_text.insert(tk.END, "WGS84 COORDINATES:\n", "header")
            self.result_text.insert(tk.END, f"  Latitude:  {data2['Latitude']:.8f}°\n")
            self.result_text.insert(tk.END, f"  Longitude: {data2['Longitude']:.8f}°\n")
            self.result_text.insert(tk.END, f"  Ellip. Ht: {data2['Ellipsoidal_Height']:.3f} m\n")
            if 'Antenna_Height' in display_data2:
                self.result_text.insert(tk.END, f"  Antenna Ht: {display_data2['Antenna_Height']:.3f} m\n")
            
            # UTM Coordinates section
            self.result_text.insert(tk.END, f"\nUTM COORDINATES (EPSG:{epsg_code}):\n", "header")
            self.result_text.insert(tk.END, f"  Easting:  {xy2[0]:.3f} m\n")
            self.result_text.insert(tk.END, f"  Northing: {xy2[1]:.3f} m\n\n")
            
            # Divider before measurement results
            self.result_text.insert(tk.END, "=" * 40 + "\n\n")
            
            # Results section with measurements
            self.result_text.insert(tk.END, "MEASUREMENT RESULTS\n", "header")
            self.result_text.insert(tk.END, "------------------\n")
            
            # Measurement results with clear formatting
            self.result_text.insert(tk.END, f"Horizontal Distance:          {xy_distance:.3f} meters\n", "result")
            self.result_text.insert(tk.END, f"Ellipsoidal Height Difference: {height_diff:.3f} meters\n", "result")
            
            # Antenna height difference if available
            if 'Antenna_Height' in data1 and 'Antenna_Height' in data2:
                antenna_diff = abs(data1['Antenna_Height'] - data2['Antenna_Height'])
                self.result_text.insert(tk.END, f"Antenna Height Difference:    {antenna_diff:.3f} meters\n", "result")
            
            # Add timestamp
            self.result_text.insert(tk.END, "\n" + "=" * 40 + "\n")
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            self.result_text.insert(tk.END, f"Analysis completed at: {timestamp}\n", "info")

            self.result_text.see(1.0)  # Scroll to top
            self.status_var.set("Calculation complete")
            self.progress_var.set("✅ Complete")
        except Exception as e:
            import traceback
            self.result_text.insert(tk.END, f"Error during computation: {str(e)}\n", "error")
            self.result_text.insert(tk.END, traceback.format_exc())
            self.status_var.set("Error in calculation")
            self.progress_var.set("❌ Error")
        
        # Re-enable the button
        self.compute_btn.config(state=tk.NORMAL)


# Main program entry
if __name__ == "__main__":
    try:
        # Create main window
        root = tk.Tk()
        root.geometry("800x600")  # Set appropriate window size
        root.minsize(640, 480)    # Set minimum window size
        
        # Create and start the application
        app = App(root)
        
        # Start main loop
        root.mainloop()
    except Exception as e:
        import traceback
        print(f"Fatal error in main program: {str(e)}")
        traceback.print_exc()
        
        # Try to show error message box
        try:
            if 'root' in locals() and root:
                messagebox.showerror("Fatal Error", f"Application encountered a fatal error: {str(e)}")
            else:
                # Create new Tk instance to show error
                error_root = tk.Tk()
                error_root.withdraw()
                messagebox.showerror("Fatal Error", f"Application failed to start: {str(e)}")
                error_root.destroy()
        except:
            # If even error dialog fails, just exit
            pass