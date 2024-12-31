import os
import sys
import math
import win32gui
import win32con
import win32com.client
from PyQt6 import QtWidgets, QtCore
from PyQt6.QtWidgets import QApplication, QMainWindow, QStatusBar, QDialog, QToolBar, QLabel, QPushButton, QVBoxLayout, QWidget, QSpinBox, QTextEdit
from PyQt6.QtGui import QIcon, QDesktopServices, QAction
from pyautocad import Autocad, APoint
from PyQt6.QtCore import QUrl, Qt


def is_autocad_running():
    try:
        # Check if AutoCAD is running using COM
        acad = win32com.client.GetObject(None, "AutoCAD.Application")
        if acad:
            return True
    except Exception as e:
        # AutoCAD is not running or some other issue occurred
        return False


def check_autocad_window():
    # Check for AutoCAD windows
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            if 'AutoCAD' in win32gui.GetWindowText(hwnd):
                windows.append(hwnd)

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return len(windows) > 0


def resource_path(relative_path):
    """ Get absolute path to resource, works for development and for PyInstaller bundle """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


app_icon_path = resource_path('app.ico')
darkMode_icon_path = resource_path('dark.png')
lightMode_icon_path = resource_path('light.png')
email_icon_path = resource_path('email.png')
license_icon_path = resource_path('license.png')


class AutoCADDividerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(QtCore.Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowIcon(QIcon(app_icon_path))
        self.setWindowTitle("AutoCAD Shape Divider")
        self.setGeometry(100, 100, 500, 500)

        self.is_dark_mode = False
        self.is_light_mode = True

        # Create the status bar
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

        # Add permanent widget to the status bar
        self.status_label = QLabel("Yousef Sedik©2024", self)
        self.status_bar.addPermanentWidget(self.status_label)

        tool_bar = QToolBar('Main Toolbar', self)
        self.addToolBar(tool_bar)

        tool_bar.addSeparator()
        self.actionButtonEmail = QAction(QIcon(email_icon_path), "Contact Me", self)
        self.actionButtonEmail.triggered.connect(self.send_email)
        tool_bar.addAction(self.actionButtonEmail)

        tool_bar.addSeparator()
        self.actionButtonLicense = QAction(QIcon(license_icon_path), 'License', self)
        self.actionButtonLicense.triggered.connect(self.show_popup)
        tool_bar.addAction(self.actionButtonLicense)

        self.label = QLabel("AutoCAD Shape Divider", self)
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.divisions_label = QLabel("Number of Divisions:", self)
        self.divisions_input = QSpinBox(self)
        self.divisions_input.setRange(2, 1000)
        self.divisions_input.setValue(20)

        self.output_log = QTextEdit(self)
        self.output_log.setReadOnly(True)
        self.output_log.setPlaceholderText("Execution logs...")

        self.select_button = QPushButton("Select Objects", self)
        self.select_button.clicked.connect(self.select_objects)

        self.execute_button = QPushButton("Execute Division", self)
        self.execute_button.clicked.connect(self.divide_shapes)

        self.clear_output_log_button = QPushButton("Clear Execution logs", self)
        self.clear_output_log_button.clicked.connect(self.clear_output_log)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.divisions_label)
        layout.addWidget(self.divisions_input)
        layout.addWidget(self.select_button)
        layout.addWidget(self.execute_button)
        layout.addWidget(self.clear_output_log_button)
        layout.addWidget(self.output_log)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.acad = None

        QtCore.QTimer.singleShot(0, self.initialize_autocad)

    @QtCore.pyqtSlot()
    def initialize_autocad(self):
        if is_autocad_running():
            self.acad = Autocad()
            if not self.acad:
                self.status_bar.showMessage("Failed to connect to AutoCAD.")
            else:
                self.status_bar.showMessage("Connected to AutoCAD.")
                self.bring_autocad_to_foreground()
        else:
            self.status_bar.showMessage("AutoCAD is not running.\nPlease start AutoCAD and try again.")

    def bring_autocad_to_foreground(self):
        try:
            hwnd = self.get_autocad_hwnd()
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Restore AutoCAD if minimized
                win32gui.SetForegroundWindow(hwnd)
                win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_TOPMOST,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                )
                return True
            else:
                return False
        except Exception as e:
            self.output_log.append(f"Error bringing AutoCAD window to the front: {e}")
            return False

    def get_autocad_hwnd(self):
        hwnds = []

        def enum_windows_callback(hwnd, lParam):
            class_name = win32gui.GetClassName(hwnd)
            if 'AutoCAD' in class_name:
                hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(enum_windows_callback, None)
        if hwnds:
            return hwnds[0]
        else:
            return None

    def clear_output_log(self):
        self.output_log.clear()

    def select_objects(self):

        self.clear_output_log()

        if self.acad is None:
            self.status_bar.showMessage("AutoCAD is not running. Please start AutoCAD.")
            return

        try:
            self.selection = self.acad.get_selection("Select one or more objects:")
            if not self.selection:
                self.output_log.append("No objects were selected.")
            else:
                self.output_log.append(f"{self.selection.Count} object(s) selected.")
        except Exception as e:
            self.output_log.append(f"Error selecting objects:\n1. AutoCAD might not be running.\n2. AutoCAD window might be minimized.\n3. Close any popup windows (if found) and try again.")

    def divide_shapes(self):
        try:
            self.clear_output_log()

            if self.acad is None:
                self.status_bar.showMessage("AutoCAD is not running. Please start AutoCAD.")
                return

            divisions = self.divisions_input.value()
            if hasattr(self, 'selection') and self.selection:
                for obj in self.selection:
                    self.divide_shape_and_connect(obj, divisions)
                self.output_log.append(f"Successfully divided shapes into {divisions} segments.")
            else:
                self.output_log.append("No objects selected.")
        except Exception as e:
            self.output_log.append(f"Error dividing the shape(s):\n1. Make sure that you have at least 1 shape selected\n2. Make sure AutoCAD interface is accessible.")

    def divide_shape_and_connect(self, obj, divisions):
        points = []
        try:
            # Divide line
            if obj.ObjectName == "AcDbLine":
                start_point = APoint(obj.StartPoint)
                end_point = APoint(obj.EndPoint)
                points = self.get_division_points_line(start_point, end_point, divisions)

            # Divide circle
            elif obj.ObjectName == "AcDbCircle":
                center_point = APoint(obj.Center)
                radius = obj.Radius
                points = self.get_division_points_circle(center_point, radius, divisions)

            # Divide polyline
            elif obj.ObjectName == "AcDbPolyline":
                points = self.get_division_points_polyline(obj, divisions)

            # Divide arc
            elif obj.ObjectName == "AcDbArc":
                center_point = APoint(obj.Center)
                radius = obj.Radius
                start_angle = obj.StartAngle
                end_angle = obj.EndAngle
                points = self.get_division_points_arc(center_point, radius, start_angle, end_angle, divisions)

            # Divide ellipse
            elif obj.ObjectName == "AcDbEllipse":
                center_point = APoint(obj.Center)
                major_axis = APoint(obj.MajorAxis)  # Ensure this is the end point of the major axis
                radius_ratio = obj.RadiusRatio
                points = self.get_division_points_ellipse(center_point, major_axis, radius_ratio, divisions)

            # Divide spline
            elif obj.ObjectName == "AcDbSpline":
                points = self.get_division_points_spline(obj, divisions)

            else:
                self.output_log.append("The selected shape is not defined in the database... yet!")
                return

            obj.Delete()
            for i in range(len(points) - 1):
                start_point = points[i]
                end_point = points[i + 1]
                self.acad.model.AddLine(start_point, end_point)

        except Exception as e:
            self.output_log.append(f"Error dividing the shape: {e}")

    def get_division_points_line(self, start_point, end_point, divisions):
        step_x = (end_point.x - start_point.x) / divisions
        step_y = (end_point.y - start_point.y) / divisions
        step_z = (end_point.z - start_point.z) / divisions

        points = [start_point]
        for i in range(1, divisions):
            new_point = APoint(
                start_point.x + step_x * i,
                start_point.y + step_y * i,
                start_point.z + step_z * i
            )
            points.append(new_point)
        points.append(end_point)
        return points

    def get_division_points_circle(self, center_point, radius, divisions):
        points = []
        for i in range(divisions):
            angle = (2 * math.pi * i) / divisions
            x = center_point.x + radius * math.cos(angle)
            y = center_point.y + radius * math.sin(angle)
            points.append(APoint(x, y, 0))
        points.append(points[0])
        return points

    def get_division_points_arc(self, center_point, radius, start_angle, end_angle, divisions):
        points = []
        angle_step = (end_angle - start_angle) / divisions
        for i in range(divisions + 1):
            angle = start_angle + angle_step * i
            x = center_point.x + radius * math.cos(angle)
            y = center_point.y + radius * math.sin(angle)
            points.append(APoint(x, y, 0))
        return points

    def get_division_points_ellipse(self, center_point, major_axis_point, radius_ratio, divisions):
        points = []
        major_radius = center_point.distance_to(major_axis_point)  # Distance between center and end of major axis
        minor_radius = major_radius * radius_ratio
        angle_step = 2 * math.pi / divisions

        for i in range(divisions):
            angle = angle_step * i
            x = center_point.x + major_radius * math.cos(angle)
            y = center_point.y + minor_radius * math.sin(angle)
            points.append(APoint(x, y, 0))
        points.append(points[0])  # Close the ellipse
        return points

    def get_division_points_polyline(self, polyline, divisions):
        points = []
        segments = polyline.NumberOfVertices - 1
        total_length = 0.0
        for i in range(segments):
            start_point = APoint(polyline.GetPointAt(i))
            end_point = APoint(polyline.GetPointAt(i + 1))
            segment_length = start_point.distance_to(end_point)
            total_length += segment_length

        step = total_length / divisions
        current_distance = 0.0

        for i in range(divisions + 1):
            distance = step * i
            accumulated_length = 0.0
            for j in range(segments):
                start_point = APoint(polyline.GetPointAt(j))
                end_point = APoint(polyline.GetPointAt(j + 1))
                segment_length = start_point.distance_to(end_point)

                if accumulated_length + segment_length >= distance:
                    segment_fraction = (distance - accumulated_length) / segment_length
                    x = start_point.x + segment_fraction * (end_point.x - start_point.x)
                    y = start_point.y + segment_fraction * (end_point.y - start_point.y)
                    points.append(APoint(x, y, 0))
                    break

                accumulated_length += segment_length

        return points

    def get_division_points_spline(self, spline, divisions):
        points = []
        num_points = spline.NumberOfControlPoints
        total_length = 0.0

        for i in range(num_points - 1):
            start_point = APoint(spline.GetPointAt(i))
            end_point = APoint(spline.GetPointAt(i + 1))
            segment_length = start_point.distance_to(end_point)
            total_length += segment_length

        step = total_length / divisions
        current_distance = 0.0

        for i in range(divisions + 1):
            distance = step * i
            accumulated_length = 0.0
            for j in range(num_points - 1):
                start_point = APoint(spline.GetPointAt(j))
                end_point = APoint(spline.GetPointAt(j + 1))
                segment_length = start_point.distance_to(end_point)

                if accumulated_length + segment_length >= distance:
                    segment_fraction = (distance - accumulated_length) / segment_length
                    x = start_point.x + segment_fraction * (end_point.x - start_point.x)
                    y = start_point.y + segment_fraction * (end_point.y - start_point.y)
                    points.append(APoint(x, y, 0))
                    break

                accumulated_length += segment_length

        return points

    def show_popup(self) -> None:
        """Show a license popup."""
        dialog = QDialog(self)
        dialog.setWindowTitle("License")
        layout = QVBoxLayout(dialog)
        text_edit = QTextEdit(dialog)
        text_edit.setText("""Software License Agreement

License Grant ->
This License Agreement (the "Agreement") is made and entered into by and between Yousef Sedik ("Author") (collectively referred to as "Licensor") and the licensee ("Licensee") who obtains the software application ("Software"). The Software is a proprietary application designed to convert curves into line segments in AutoCAD.

Use of Software ->
The Licensee is granted a non-exclusive, non-transferable license to use the Software solely for the Licensee's internal engineering purposes. The Software is provided free of charge and may not be used for any commercial or governmental purposes without obtaining a separate license.

Restrictions ->
The Licensee may not:
Reverse engineer, decompile, disassemble, or otherwise attempt to discover the source code of the Software.
Use the Software in any manner that violates the terms of this Agreement.
Ownership

But the Licensee may:
Copy, or distribute the Software without prior written consent from the Licensor.

The Software and any associated documentation, including any updates, are the intellectual property of the Licensor. The Licensee does not acquire any ownership rights in the Software by virtue of this Agreement.

Termination ->
The Licensor reserve the right to terminate this Agreement and the License granted hereunder if the Licensee breaches any terms of this Agreement. Upon termination, the Licensee must cease all use of the Software and destroy all copies in their possession.

Warranty Disclaimer ->
The Software is provided "as-is" without any warranties, express or implied. The Licensor make no representations or warranties regarding the Software's performance or its fitness for a particular purpose.

Limitation of Liability ->
In no event shall the Licensor be liable for any damages arising from the use or inability to use the Software, including but not limited to incidental or consequential damages.

Governing Law ->
This Agreement shall be governed by and construed in accordance with the laws of Egypt, without regard to its conflict of laws principles.
Entire Agreement

This Agreement constitutes the entire agreement between the Licensor and the Licensee concerning the Software and supersedes all prior agreements or understandings, whether written or oral, relating to the Software.

By using the Software, the Licensee acknowledges that they have read, understood, and agree to be bound by the terms and conditions of this Agreement.

Licensor:
Yousef Sedik
Email: yousefsedik.bus@gmail.com""")
        text_edit.setReadOnly(True)
        layout.addWidget(text_edit)
        dialog.exec()

    @staticmethod
    def send_email() -> None:
        """Open the mail app on the user's pc with Mail to, CC and Subject filled automatically"""
        email_address = "yousefsedik.bus@gmail.com"
        subject = "Regarding AutoCad Curve Divider V2.0"
        body = ""
        mailto_link = f"mailto:{email_address}?&subject={subject}&body={body}"

        QDesktopServices.openUrl(QUrl(mailto_link))  # type: ignore


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutoCADDividerApp()
    window.show()
    app.exec()
