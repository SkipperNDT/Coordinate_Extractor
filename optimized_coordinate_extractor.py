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
        master.title("Coordinate Distance Calculator")

        # Status bar for loading information
        self.status_var = tk.StringVar()
        self.status_var.set("Initializing...")
        self.status_bar = tk.Label(master, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Main frame
        self.main_frame = tk.Frame(master)
        self.main_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.image1_path = ""
        self.image2_path = ""

        # Buttons frame
        self.btn_frame = tk.Frame(self.main_frame)
        self.btn_frame.pack(pady=5)

        self.btn1 = tk.Button(self.btn_frame, text="Select Image 1", command=self.load_image1, width=15)
        self.btn1.pack(side=tk.LEFT, padx=5)

        self.btn2 = tk.Button(self.btn_frame, text="Select Image 2", command=self.load_image2, width=15)
        self.btn2.pack(side=tk.LEFT, padx=5)

        self.compute_btn = tk.Button(self.btn_frame, text="Compute Distance", command=self.compute_distance, width=15)
        self.compute_btn.pack(side=tk.LEFT, padx=5)
        self.compute_btn.config(state=tk.DISABLED)  # Initially disabled

        # Text box for results
        self.result_frame = tk.Frame(self.main_frame)
        self.result_frame.pack(pady=5, fill=tk.BOTH, expand=True)
        
        self.result_text = tk.Text(self.result_frame, height=20, width=80)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbar
        self.scrollbar = tk.Scrollbar(self.result_frame, command=self.result_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=self.scrollbar.set)

        # Initial message
        self.result_text.insert(tk.END, "Welcome to Coordinate Distance Calculator!\n")
        self.result_text.insert(tk.END, "Please select two images containing coordinate information.\n\n")

        # Start background loading of modules
        self.load_modules_thread = threading.Thread(target=self.load_modules_in_background)
        self.load_modules_thread.daemon = True
        self.load_modules_thread.start()

    def load_modules_in_background(self):
        """Load heavy modules in background to improve startup time"""
        global TESSERACT_LOADED, PIL_LOADED, PYPROJ_LOADED, NUMPY_LOADED
        
        self.status_var.set("Loading PIL (Image processing)...")
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
            # Enable compute button if both images are selected
            if self.image1_path and self.image2_path:
                self.compute_btn.config(state=tk.NORMAL)
        else:
            self.status_var.set("Some components failed to load")

    def load_image1(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.tif")])
        if path:
            self.image1_path = path
            self.result_text.insert(tk.END, f"Selected Image 1: {path}\n")
            self.result_text.see(tk.END)
            # Enable compute button if both images are selected and modules loaded
            if self.image2_path and TESSERACT_LOADED and PYPROJ_LOADED:
                self.compute_btn.config(state=tk.NORMAL)

    def load_image2(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.tif")])
        if path:
            self.image2_path = path
            self.result_text.insert(tk.END, f"Selected Image 2: {path}\n")
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
        # Latitude
        lat_match = re.search(r"Latitude\s+([0-9.,°]+)", text)
        if lat_match:
            try:
                result['Latitude'] = float(lat_match.group(1).replace(",", ".").replace("°", ""))
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Latitude: {str(e)}\n")
        else:
            self.result_text.insert(tk.END, "Latitude not found in OCR text\n")

        # Longitude
        lon_match = re.search(r"Longitude\s+([0-9.,°]+)", text)
        if lon_match:
            try:
                result['Longitude'] = float(lon_match.group(1).replace(",", ".").replace("°", ""))
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Longitude: {str(e)}\n")
        else:
            self.result_text.insert(tk.END, "Longitude not found in OCR text\n")

        # Ellipsoidal Height
        ellip_match = re.search(r"Hauteur ellipsoïdale\s+([\d.,]+)\s*m", text)
        if ellip_match:
            try:
                result['Ellipsoidal_Height'] = float(ellip_match.group(1).replace(",", "."))
            except ValueError as e:
                self.result_text.insert(tk.END, f"Error parsing Ellipsoidal Height: {str(e)}\n")
        else:
            self.result_text.insert(tk.END, "Ellipsoidal Height not found in OCR text\n")

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
            self.result_text.insert(tk.END, f"Using Tesseract: {self.pytesseract.pytesseract.tesseract_cmd}\n\n")
        else:
            self.result_text.insert(tk.END, "ERROR: Tesseract OCR not loaded\n")
            self.compute_btn.config(state=tk.NORMAL)
            return

        if not self.image1_path or not self.image2_path:
            messagebox.showerror("Error", "Please select both images first.")
            self.compute_btn.config(state=tk.NORMAL)
            return

        # Use a thread for processing to keep UI responsive
        threading.Thread(target=self._compute_distance_thread).start()

    def _compute_distance_thread(self):
        # Extract parameters
        self.status_var.set("Processing Image 1...")
        data1, raw1 = self.extract_parameters(self.image1_path)
        
        self.status_var.set("Processing Image 2...")
        data2, raw2 = self.extract_parameters(self.image2_path)

        if not data1 or not data2:
            self.result_text.insert(tk.END, "\nOCR extraction failed. Check images and text layout.\n")
            if raw1:
                self.result_text.insert(tk.END, f"\nRaw OCR Text 1:\n{raw1}\n\n")
            if raw2:
                self.result_text.insert(tk.END, f"Raw OCR Text 2:\n{raw2}\n")
            self.compute_btn.config(state=tk.NORMAL)
            self.status_var.set("Ready")
            return

        # Check if all required fields are present
        required_fields = ['Latitude', 'Longitude', 'Ellipsoidal_Height']
        missing_fields = False
        for field in required_fields:
            if field not in data1:
                self.result_text.insert(tk.END, f"Missing field in Image 1: {field}\n")
                missing_fields = True
            if field not in data2:
                self.result_text.insert(tk.END, f"Missing field in Image 2: {field}\n")
                missing_fields = True
        
        if missing_fields:
            self.result_text.insert(tk.END, f"\nImage 1 data: {data1}\n")
            self.result_text.insert(tk.END, f"Image 2 data: {data2}\n")
            self.compute_btn.config(state=tk.NORMAL)
            self.status_var.set("Ready")
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

            # Show results
            result_lines = [
                "---- Extracted Parameters ----",
                f"Image 1: {data1}",
                f"Image 2: {data2}",
                "",
                f"Detected UTM EPSG Code: {epsg_code}",
                "",
                f"UTM Coordinates (meters):",
                f"Image 1: X={xy1[0]:.2f}, Y={xy1[1]:.2f}",
                f"Image 2: X={xy2[0]:.2f}, Y={xy2[1]:.2f}",
                "",
                f"Horizontal Distance (meters): {xy_distance:.3f}",
                f"Ellipsoidal Height Difference (meters): {height_diff:.2f}"
            ]

            self.result_text.insert(tk.END, "\n".join(result_lines))
            self.result_text.see(tk.END)
            self.status_var.set("Calculation complete")
        except Exception as e:
            import traceback
            self.result_text.insert(tk.END, f"Error during computation: {str(e)}\n")
            self.result_text.insert(tk.END, traceback.format_exc())
            self.status_var.set("Error in calculation")
        
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