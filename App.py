"""
UNIFIED ASSESSMENT CHART GENERATOR
Combined code from 4 sheets - Creates Radar Charts, Bar Charts, and Assessment Visualizations
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from math import pi
import matplotlib.font_manager as fm
import os
import re

# Try to import Persian text reshaping libraries
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_RESHAPER = True
    print("✓ Arabic reshaper loaded successfully")
except ImportError:
    HAS_RESHAPER = False
    print("⚠ Please install required packages:")
    print("  pip install arabic-reshaper python-bidi")
    print("Without these, Persian text will appear separated!")


# ==================== TEXT PROCESSING ====================

class PersianTextProcessor:
    """Handles Persian/Farsi text reshaping, wrapping, and formatting"""
    
    @staticmethod
    def reshape_text(text):
        """Properly reshape Persian/Farsi text for correct display"""
        if not HAS_RESHAPER or not text or pd.isna(text):
            return text
        try:
            reshaped_text = arabic_reshaper.reshape(str(text))
            return get_display(reshaped_text)
        except Exception as e:
            print(f"Error reshaping text: {e}")
            return text
    
    @staticmethod
    def wrap_text(text, max_chars=15):
        """Wrap text into multiple lines after max_chars characters"""
        if not text or pd.isna(text):
            return text
        
        text = str(text)
        if len(text) <= max_chars:
            return text
        
        words = text.split()
        lines = []
        current_line = ""
        
        for word in words:
            if len(current_line + " " + word) <= max_chars:
                if current_line:
                    current_line += " " + word
                else:
                    current_line = word
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        
        if current_line:
            lines.append(current_line)
        
        return "\n".join(lines)
    
    @staticmethod
    def process_row_name(text, max_chars=20):
        """Process row name: reshape and wrap"""
        reshaped = PersianTextProcessor.reshape_text(str(text))
        return PersianTextProcessor.wrap_text(reshaped, max_chars)
    
    @staticmethod
    def format_number(value, decimal_places=1):
        """Format number for display (removes .0 for whole numbers)"""
        if pd.isna(value):
            return "0"
        try:
            rounded = round(float(value), decimal_places)
            if abs(rounded - int(rounded)) < 0.01:
                return str(int(rounded))
            else:
                return f"{rounded:.{decimal_places}f}"
        except (ValueError, TypeError):
            return "0"
    
    @staticmethod
    def clean_filename(text):
        """Remove invalid characters from filename"""
        invalid_chars = r'[\\/*?:"<>|]'
        cleaned = re.sub(invalid_chars, "_", str(text))
        return cleaned[:100] if len(cleaned) > 100 else cleaned


class FontManager:
    """Manages font configuration for Persian text"""
    
    @staticmethod
    def set_persian_font():
        """Find and set a Persian font for matplotlib"""
        persian_fonts = ['IRANSans(FaNum)']
        for font_name in persian_fonts:
            try:
                if fm.findfont(font_name, fallback_to_default=False):
                    print(f"✓ Using Persian font: {font_name}")
                    plt.rcParams.update({'font.family': font_name, 'axes.unicode_minus': False})
                    return font_name
            except:
                continue
        print("⚠ No Persian font found. Using default font.")
        return None


# ==================== DATA LOADING ====================

class DataLoader:
    """Handles data loading and preprocessing for all chart types"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.categories = []
        self.values_matrix = None
        self.raw_data = None
        self.parents = None
        self.children = None
        self.values = None
        self.mapping = None
        
    def load_data(self, sheet_name):
        """Load data from Excel file with specified sheet"""
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
        return self.df
    
    def load_radar_data(self, sheet_name='3.1'):
        """Load and prepare data for radar charts (Sheet 3.1)"""
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
        self.df = self.df.iloc[1:-2]
        self.df['new'] = 5
        self.df = self.df.T
        
        # Extract categories from first row
        categories_row = self.df.iloc[0].values
        self.df = self.df.iloc[1:]
        self.df.columns = categories_row
        
        # Remove the 'new' column if it exists
        if 'new' in self.df.columns:
            self.df = self.df.drop(columns=['new'])
        
        self.categories = self.df.columns.tolist()
        
        # Convert all data to numeric
        for col in self.categories:
            self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
        
        # Create values matrix
        self.values_matrix = self.df[self.categories].values
        
        # Remove rows with all NaN values
        valid_rows = ~np.isnan(self.values_matrix).all(axis=1)
        if not valid_rows.all():
            print(f"\n⚠ Removing {np.sum(~valid_rows)} rows with all NaN values")
            self.values_matrix = self.values_matrix[valid_rows]
            self.df = self.df[valid_rows]
        
        return self.df, self.categories, self.values_matrix
    
    def load_bar_chart_data(self, sheet_name='3.3'):
        """Load and prepare data for bar charts (Sheet 3.3)"""
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
        self.df.iloc[1] = self.df.iloc[1].ffill()
        self.df = self.df.iloc[1:, [0, 3, 6, 9, 12]]
        return self.df.reset_index(drop=True)
    
    def load_assessment_data(self, sheet_name='3.4'):
        """Load and prepare data for assessment charts (Sheet 3.4)"""
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
        
        # Remove first 10 rows and forward fill
        self.df = self.df.iloc[10:].copy()
        self.df = self.df.ffill()
        
        # Drop empty columns and rows
        self.df = self.df.dropna(axis=1, how='all')
        self.df = self.df.dropna(axis=0, how='all')
        
        # Detect numeric columns
        numeric_columns = []
        for col in self.df.columns:
            non_null = self.df[col].dropna()
            if len(non_null) > 0:
                numeric_count = sum(pd.to_numeric(non_null, errors='coerce').notna())
                if numeric_count / len(non_null) > 0.5:
                    numeric_columns.append(col)
        
        if len(numeric_columns) >= 1:
            self.row_name_col = self.df.columns[0]
            self.value_cols = numeric_columns
            
            # Generate categories from column headers
            self.categories = []
            for col in self.value_cols:
                header_val = self.df[col].iloc[0] if len(self.df) > 0 else f"Value {col}"
                if pd.notna(header_val) and not isinstance(header_val, (int, float)):
                    self.categories.append(str(header_val))
                else:
                    self.categories.append(f"معیار {len(self.categories)+1}")
            
            # Use default Persian names if no meaningful categories found
            if len(self.categories) < len(self.value_cols) or all("معیار" in c for c in self.categories):
                default_categories = [
                    'جریان های ارزش و فرآیندها',
                    "اطلاعات و فناوری",
                    "سازمان ها و افراد",
                    "شرکاء و تأمین کنندگان"
                ]
                self.categories = default_categories[:len(self.value_cols)]
        else:
            # Fallback
            self.row_name_col = self.df.columns[0]
            self.value_cols = self.df.columns[1:5]
            self.categories = [
                'جریان های ارزش و فرآیندها',
                "اطلاعات و فناوری",
                "سازمان ها و افراد",
                "شرکاء و تأمین کنندگان"
            ][:len(self.value_cols)]
        
        return self.df, self.categories, self.value_cols
    
    def load_radar_35_data(self, sheet_name='3.5'):
        """Load and prepare data for 3.5 radar charts"""
        self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
        
        # Forward fill the parent row
        self.df.iloc[1] = self.df.iloc[1].ffill()
        
        # Extract data
        self.parents = self.df.iloc[1, 1:]
        self.children = self.df.iloc[2, 1:]
        self.values = self.df.iloc[3, 1:]
        
        # Create mapping
        self.mapping = {}
        for parent, child, value in zip(self.parents, self.children, self.values):
            self.mapping[child] = {'parent': parent, 'children': child, 'value': float(value)}
        
        # Create DataFrame
        df_processed = pd.DataFrame([{'parent': data['parent'], 'children': child, 'values': data['value']}
                                     for child, data in self.mapping.items()])
        
        df_processed['new'] = 100
        self.df = df_processed
        return self.df


# ==================== RADAR CHART (3.1) ====================

class RadarChart31:
    """Creates radar charts for Sheet 3.1 data"""
    
    def __init__(self, df, categories, values_matrix):
        self.df = df
        self.categories = categories
        self.values_matrix = values_matrix
        self.num_categories = len(categories)
        self.fig = None
        self.ax = None
        self.angles = None
        self.max_value = None
        self.min_value = None
        self.line_colors = ["#4155c6", "#ff7e79", "#2ecc71", "#f39c12", "#9b59b6", "#e74c3c"]
        self.text_processor = PersianTextProcessor()
        
    def calculate_angles(self):
        """Calculate angles for radar chart categories"""
        self.angles = np.linspace(0, 2*pi, self.num_categories, endpoint=False).tolist()
        return self.angles
    
    def calculate_value_range(self):
        """Calculate min and max values from data"""
        self.max_value = np.nanmax(self.values_matrix)
        self.min_value = np.nanmin(self.values_matrix)
        if self.max_value == 0:
            print("⚠ All values are zero! Adjusting max_value to 1 for visualization")
            self.max_value = 1
        return self.max_value, self.min_value
    
    def create_figure(self):
        """Create the matplotlib figure and axis"""
        figsize = max(12, self.num_categories)
        self.fig, self.ax = plt.subplots(figsize=(figsize, figsize), subplot_kw=dict(projection='polar'))
        return self.fig, self.ax
    
    def draw_background(self):
        """Draw background circles and grid"""
        num_circles = 5
        circle_levels = np.linspace(0, self.max_value, num_circles + 1)
        theta_circle = np.linspace(0, 2*pi, 400)
        
        for i in range(num_circles):
            color = "#ffffff" if i % 2 == 0 else "#f0f0f0"
            self.ax.fill_between(theta_circle, circle_levels[i], circle_levels[i+1], 
                                color=color, zorder=0)
        
        for level in circle_levels[1:]:
            self.ax.plot(theta_circle, [level] * len(theta_circle), 
                        color="#cccccc", linewidth=1, zorder=1)
        
        radial_angles = np.array([[angle, angle] for angle in self.angles])
        radial_values = np.array([[0, self.max_value] for _ in self.angles])
        for angle_vals, value_vals in zip(radial_angles, radial_values):
            self.ax.plot(angle_vals, value_vals, color="#cccccc", linewidth=1, 
                        zorder=1, linestyle='-')
    
    def draw_data_lines(self):
        """Draw data lines and labels for each row"""
        plot_angles = self.angles + self.angles[:1]
        offset_inside = self.max_value * 0.11
        offset_outside = self.max_value * 0.08
        
        for idx, (index_name, row_values) in enumerate(self.df.iterrows()):
            values = row_values[self.categories].values.tolist()
            values_clean = [0 if pd.isna(v) else v for v in values]
            values_closed = values_clean + values_clean[:1]
            
            color = self.line_colors[idx % len(self.line_colors)]
            label_text = self.text_processor.reshape_text(str(index_name))
            
            if any(v > 0 for v in values_clean):
                self.ax.plot(plot_angles, values_closed, 'o-', linewidth=1.5, label=label_text,
                            color=color, zorder=11, markersize=3, markeredgewidth=1.5)
                
                label_position_offset = offset_inside if idx == 1 else offset_outside
                
                for angle, value in zip(self.angles, values_clean):
                    label_position = value - label_position_offset if idx == 1 else value + label_position_offset
                    
                    if not pd.isna(value):
                        value_text = self.text_processor.format_number(value)
                        self.ax.text(angle, label_position, value_text, ha='center', va='center',
                                    fontsize=9, fontweight='bold', color=color)
    
    def configure_axes(self):
        """Configure axis properties"""
        self.ax.set_ylim(0, self.max_value)
        self.ax.yaxis.grid(False)
        self.ax.xaxis.grid(False)
        self.ax.set_yticklabels([])
        self.ax.spines['polar'].set_color("#999999")
        self.ax.spines['polar'].set_linewidth(1.5)
        self.ax.set_xticks(self.angles)
        self.ax.set_xticklabels([])
    
    def add_category_labels(self):
        """Add category labels with custom positioning"""
        label_distance = self.max_value * 1.1
        gap_distance = self.max_value * 0.01
        
        for i, (category, angle) in enumerate(zip(self.categories, self.angles)):
            category_str = str(category)
            display_text = self.text_processor.reshape_text(category_str)
            
            cos_angle = np.cos(angle)
            sign = 1 if cos_angle > 0 else -1
            ha = 'left' if cos_angle > 0 else 'right'
            
            x_pos = label_distance * cos_angle + sign * gap_distance
            y_pos = label_distance * np.sin(angle)
            
            new_radius = np.hypot(x_pos, y_pos)
            new_angle = np.arctan2(y_pos, x_pos)
            if new_angle < 0:
                new_angle += 2*pi
            
            self.ax.plot([angle, new_angle], [self.max_value, new_radius * 0.98], 
                        color='gray', linewidth=1, linestyle='--', alpha=0.5, zorder=5)
            
            self.ax.text(new_angle, new_radius, display_text, ha=ha, va='center',
                        fontsize=11, rotation=0,
                        bbox=dict(boxstyle='round,pad=0.4', facecolor='white', 
                                 alpha=0.95, edgecolor='white', linewidth=0.8))
    
    def draw(self):
        """Draw the complete radar chart"""
        self.calculate_angles()
        self.calculate_value_range()
        self.create_figure()
        self.draw_background()
        self.draw_data_lines()
        self.configure_axes()
        self.add_category_labels()
        return self.fig, self.ax
    
    def save_chart(self, output_folder, filename='radar_chart_31.png'):
        """Save the chart to a file"""
        os.makedirs(output_folder, exist_ok=True)
        save_path = os.path.join(output_folder, filename)
        self.fig.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
        print(f"✓ Radar chart saved to: {save_path}")
        return save_path
    
    def show(self):
        """Display the chart"""
        plt.show()


# ==================== BAR CHART (3.3) ====================

class BarChartGenerator33:
    """Generates bar charts for Sheet 3.3 data"""
    
    def __init__(self, data_loader):
        self.data_loader = data_loader
        self.text_processor = PersianTextProcessor()
        self.categories = [
            'جریان های ارزش و فرآیندها',
            "اطلاعات و فناوری",
            "سازمان ها و افراد",
            "شرکاء و تأمین کنندگان"
        ]
        self.colors = ['#298091', '#38AAC1', '#40C1DD', '#64E6FC']
        self.output_folder = None
        
    def setup_output_folder(self):
        """Create output directory on Desktop"""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.output_folder = os.path.join(desktop_path, "Assessment_Charts_3.3")
        os.makedirs(self.output_folder, exist_ok=True)
        print(f"✓ Charts will be saved to: {self.output_folder}")
        return self.output_folder
    
    def generate_chart(self, row_data, row_idx, config=None):
        """Generate a single bar chart from row data"""
        config = config or {}
        chart_width = config.get('chart_width', 6)
        chart_height = config.get('chart_height', 4)
        bar_width = config.get('bar_width', 0.9)
        decimal_places = config.get('decimal_places', 1)
        
        # Extract row name and values
        row_name = row_data.iloc[0]
        raw_values = row_data.iloc[1:5].tolist()
        
        # Process values
        values = [round(float(v), decimal_places) if pd.notna(v) else 0 for v in raw_values]
        
        # Process text
        processed_name = self.text_processor.process_row_name(row_name, max_chars=20)
        
        # Get reshaped categories
        reshaped_categories = [self.text_processor.reshape_text(cat) for cat in self.categories]
        
        # Create chart
        fig, ax = plt.subplots(figsize=(chart_width, chart_height))
        
        # Create bars
        bars = ax.bar(reshaped_categories, values, color=self.colors,
                     edgecolor='white', linewidth=2, width=bar_width)
        
        # Setup axes
        ax.set_facecolor('white')
        ax.set_xlim(-0.5, 3.5)
        max_value = max(values) if values else 100
        y_max = min(105, max_value + 10) if max_value < 95 else 105
        ax.set_ylim(0, y_max)
        ax.set_yticks([0, 20, 40, 60, 80, 100])
        ax.set_yticklabels(['0', '20', '40', '60', '80', '100'])
        
        ax.grid(axis='y', color='#ECECEC', linestyle='-', linewidth=1.5)
        ax.set_axisbelow(True)
        ax.grid(axis='x', visible=False)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#979797')
        ax.spines['bottom'].set_color('#979797')
        ax.spines['left'].set_linewidth(1.5)
        ax.spines['bottom'].set_linewidth(1.5)
        
        # Add value labels
        for bar, value in zip(bars, values):
            display_value = self.text_processor.format_number(value, decimal_places)
            if value >= 15:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() - 8,
                       display_value, ha='center', va='top', fontweight='bold',
                       fontsize=11, color='white')
            else:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3,
                       display_value, ha='center', va='bottom', fontweight='bold',
                       fontsize=11, color='#298091')
        
        plt.xticks(ha='center')
        plt.tight_layout()
        
        # Save chart
        if config.get('save_charts', True):
            clean_name = self.text_processor.clean_filename(row_name)
            filename = f"Row_{row_idx+1}_{clean_name}.{config.get('save_format', 'png')}"
            filepath = os.path.join(self.output_folder, filename)
            plt.savefig(filepath, dpi=config.get('dpi', 300), bbox_inches='tight', facecolor='white')
            print(f"✓ Chart saved: {filename}")
        
        plt.close(fig)
        return processed_name, values
    
    def generate_all_charts(self, config=None):
        """Generate charts for all rows"""
        config = config or {}
        config.setdefault('save_charts', True)
        
        if self.output_folder is None:
            self.setup_output_folder()
        
        reshaped_data = self.data_loader
        print("\n" + "="*60)
        print("Starting chart generation (3.3)...")
        print("="*60)
        
        for row_idx in range(1, len(reshaped_data)):
            row_data = reshaped_data.iloc[row_idx]
            self.generate_chart(row_data, row_idx, config)
        
        print("\n✅ All charts generated successfully!")


# ==================== ASSESSMENT CHART (3.4) ====================

class AssessmentChartGenerator34:
    """Generates assessment bar charts for Sheet 3.4 data"""
    
    def __init__(self, data_loader, categories, value_cols):
        self.data_loader = data_loader
        self.categories = categories
        self.value_cols = value_cols
        self.text_processor = PersianTextProcessor()
        self.colors = ['#298091', '#38AAC1', '#40C1DD', '#64E6FC']
        self.output_folder = None
        
    def setup_output_folder(self):
        """Create output directory on Desktop"""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.output_folder = os.path.join(desktop_path, "Assessment_Charts_3.4")
        os.makedirs(self.output_folder, exist_ok=True)
        print(f"✓ Charts will be saved to: {self.output_folder}")
        return self.output_folder
    
    def generate_all_charts(self, config=None):
        """Generate charts for all rows"""
        config = config or {}
        config.setdefault('multiply_by_100', True)
        config.setdefault('save_charts', True)
        config.setdefault('chart_width', 6)
        config.setdefault('chart_height', 4)
        config.setdefault('bar_width', 0.8)
        config.setdefault('decimal_places', 1)
        
        if config.get('save_charts') and self.output_folder is None:
            self.setup_output_folder()
        
        # Set Persian font
        FontManager.set_persian_font()
        
        print("\n" + "="*50)
        print("Creating assessment charts (3.4)...")
        print("="*50)
        
        df_reset = self.data_loader.reset_index(drop=True)
        reshaped_categories = [self.text_processor.reshape_text(cat) for cat in self.categories]
        charts_created = 0
        
        for row_idx in range(len(df_reset)):
            row = df_reset.iloc[row_idx]
            row_name = row[self.data_loader.columns[0]]
            
            # Extract values
            values = []
            for col in self.value_cols:
                val = row[col]
                if pd.notna(val):
                    val = float(val)
                    if config.get('multiply_by_100'):
                        val = val * 100
                    values.append(round(val, config['decimal_places']))
                else:
                    values.append(0)
            
            if all(v == 0 for v in values):
                print(f"Row {row_idx + 1}: '{row_name}' - All values zero, skipping...")
                continue
            
            # Create chart
            self._create_single_chart(row_idx, row_name, values, reshaped_categories, config)
            charts_created += 1
        
        print(f"\n✅ Chart generation complete! Charts created: {charts_created}")
        return charts_created
    
    def _create_single_chart(self, row_idx, row_name, values, reshaped_categories, config):
        """Create a single assessment bar chart"""
        reshaped_row_name = self.text_processor.reshape_text(str(row_name))
        
        fig, ax = plt.subplots(figsize=(config['chart_width'], config['chart_height']))
        
        bars = ax.bar(reshaped_categories, values, 
                     color=self.colors[:len(values)],
                     edgecolor='white', linewidth=2, 
                     width=config['bar_width'])
        
        ax.set_facecolor('white')
        ax.set_xlim(-0.5, len(self.categories) - 0.5)
        
        max_value = max(values) if values else 100
        y_max = min(105, max_value + 10) if max_value < 95 else 105
        ax.set_ylim(0, y_max)
        ax.set_yticks([0, 20, 40, 60, 80, 100])
        ax.set_yticklabels(['0', '20', '40', '60', '80', '100'])
        
        ax.grid(axis='y', color='#ECECEC', linestyle='-', linewidth=1.5)
        ax.set_axisbelow(True)
        ax.grid(axis='x', visible=False)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#979797')
        ax.spines['bottom'].set_color('#979797')
        ax.spines['left'].set_linewidth(1.5)
        ax.spines['bottom'].set_linewidth(1.5)
        
        for bar, value in zip(bars, values):
            display_value = self.text_processor.format_number(value, config['decimal_places'])
            if value >= 15:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() - 8,
                       display_value, ha='center', va='top', 
                       fontweight='bold', fontsize=11, color='white')
            else:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3,
                       display_value, ha='center', va='bottom', 
                       fontweight='bold', fontsize=11, color='#298091')
        
        plt.xticks(ha='center', rotation=0)
        plt.tight_layout()
        
        if config.get('save_charts'):
            clean_name = self.text_processor.clean_filename(row_name)
            filename = f"Row_{row_idx+1}_{clean_name}.{config.get('save_format', 'png')}"
            filepath = os.path.join(self.output_folder, filename)
            plt.savefig(filepath, dpi=config.get('dpi', 300), bbox_inches='tight', facecolor='white')
            print(f"✓ Chart saved: {filename}")
        
        plt.close(fig)


# ==================== RADAR CHART (3.5) ====================

class RadarChart35:
    """Creates radar charts for Sheet 3.5 data"""
    
    def __init__(self, parent_name, df, text_processor, max_chars_per_line=50):
        self.parent_name = parent_name
        self.df = df
        self.text_processor = text_processor
        self.max_chars_per_line = max_chars_per_line
        self.fig = None
        self.ax = None
        
    def prepare_data(self):
        """Prepare data for plotting"""
        current_df = self.df[self.df['parent'] == self.parent_name]
        headers = list(current_df.columns)
        categories = current_df[headers[1]].tolist()
        numeric_columns = current_df.columns[2:]
        df_for_plot = current_df.set_index(headers[0])[numeric_columns].T
        df_for_plot = df_for_plot.apply(pd.to_numeric, errors='coerce')
        return headers, categories, df_for_plot, current_df
    
    def create_chart(self):
        """Create the radar chart"""
        headers, categories, df_for_plot, current_df = self.prepare_data()
        
        num_categories = len(categories)
        values_matrix = df_for_plot.values
        max_value = values_matrix.max()
        angles = np.linspace(0, 2*pi, num_categories, endpoint=False).tolist()
        show_percentage = (max_value == 100)
        
        figsize = max(12, num_categories * 1.8)
        fig, ax = plt.subplots(figsize=(figsize, figsize), subplot_kw=dict(projection='polar'))
        self.fig, self.ax = fig, ax
        
        # Draw background circles
        num_circles = 5
        circle_levels = np.linspace(0, max_value, num_circles + 1)
        theta_circle = np.linspace(0, 2*pi, 400)
        
        for j in range(num_circles):
            color = "#ffffff" if j % 2 == 0 else "#f0f0f0"
            ax.fill_between(theta_circle, circle_levels[j], circle_levels[j+1], 
                            color=color, zorder=0)
        
        for level in circle_levels[1:]:
            ax.plot(theta_circle, [level] * len(theta_circle), 
                    color="#cccccc", linewidth=1, zorder=1)
        
        radial_angles = np.array([[angle, angle] for angle in angles])
        radial_values = np.array([[0, max_value] for _ in angles])
        for angle_vals, value_vals in zip(radial_angles, radial_values):
            ax.plot(angle_vals, value_vals, color="#cccccc", linewidth=1, 
                    zorder=1, linestyle='-')
        
        # Plot data lines
        line_colors = ["#4155c6", "#ff7e79", "#2ecc71", "#f39c12", "#9b59b6", "#e74c3c"]
        plot_angles = angles + angles[:1]
        offset_inside = max_value * 0.11
        offset_outside = max_value * 0.08
        
        for idx, (index_name, row_values) in enumerate(df_for_plot.iterrows()):
            values = row_values.values.tolist()
            values_closed = values + values[:1]
            
            color = line_colors[idx % len(line_colors)]
            label_text = self.text_processor.reshape_text(str(index_name))
            
            ax.plot(plot_angles, values_closed, 'o-', linewidth=1.5, color=color, 
                   zorder=11, markersize=3, markeredgewidth=1.5)
            
            label_position_offset = offset_inside if idx == 1 else offset_outside
            
            for angle, value in zip(angles, values):
                label_position = value - label_position_offset if idx == 1 else value + label_position_offset
                value_text = self.text_processor.format_number(value)
                if show_percentage:
                    value_text += '%'
                ax.text(angle, label_position, value_text, ha='center', va='center',
                        fontsize=9, fontweight='bold', color=color)
        
        ax.set_ylim(0, max_value)
        ax.yaxis.grid(False)
        ax.xaxis.grid(False)
        ax.set_yticklabels([])
        ax.spines['polar'].set_color("#999999")
        ax.spines['polar'].set_linewidth(1.5)
        ax.set_xticks(angles)
        ax.set_xticklabels([])
        
        # Add category labels
        self._add_category_labels(categories, angles, max_value)
        
        return fig, ax
    
    def _add_category_labels(self, categories, angles, max_value):
        """Add category labels with custom positioning"""
        label_distance = max_value * 1.1
        gap_distance = max_value * 0.01
        
        for j, (category, angle) in enumerate(zip(categories, angles)):
            wrapped_category = self.text_processor.wrap_text(str(category), self.max_chars_per_line)
            display_text = self.text_processor.reshape_text(wrapped_category)
            
            cos_angle = np.cos(angle)
            sign = 1 if cos_angle > 0 else -1
            ha = 'left' if cos_angle > 0 else 'right'
            
            x_pos = label_distance * cos_angle + sign * gap_distance
            y_pos = label_distance * np.sin(angle)
            
            new_radius = np.hypot(x_pos, y_pos)
            new_angle = np.arctan2(y_pos, x_pos)
            if new_angle < 0:
                new_angle += 2*pi
            
            self.ax.plot([angle, new_angle], [max_value, new_radius * 0.98], 
                        color='gray', linewidth=1, linestyle='--', alpha=0.5, zorder=5)
            
            self.ax.text(new_angle, new_radius, display_text, ha=ha, va='center',
                        fontsize=11, rotation=0, linespacing=1.2,
                        bbox=dict(boxstyle='round,pad=0.4', facecolor='white', 
                                 alpha=0.95, edgecolor='white', linewidth=0.8))
    
    def save_chart(self, output_folder, filename):
        """Save the chart to file"""
        if self.fig:
            save_path = os.path.join(output_folder, filename)
            self.fig.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
            return save_path
        return None


class RadarChartGenerator35:
    """Main class to generate multiple radar charts for 3.5"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.data_loader = None
        self.df = None
        self.text_processor = PersianTextProcessor()
        self.output_folder = None
        self.unique_parents = None
        
    def load_and_preprocess(self):
        """Load and preprocess the data"""
        self.data_loader = DataLoader(self.file_path)
        self.df = self.data_loader.load_radar_35_data(sheet_name='3.5')
        self.unique_parents = self.df['parent'].unique()
        return self.df
    
    def setup_output_folder(self):
        """Create output folder for saving charts"""
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.output_folder = os.path.join(desktop_path, "Assessment_Charts_3.5")
        os.makedirs(self.output_folder, exist_ok=True)
        print(f"\n✓ Charts will be saved to: {self.output_folder}")
        return self.output_folder
    
    def generate_all_charts(self):
        """Generate all radar charts"""
        self.setup_output_folder()
        FontManager.set_persian_font()
        
        mapping_file = os.path.join(self.output_folder, 'chart_mapping.txt')
        with open(mapping_file, 'w', encoding='utf-8') as f:
            f.write("Chart Filename - Parent Name (Persian)\n")
            f.write("=" * 50 + "\n")
        
        for i, parent_name in enumerate(self.unique_parents, 1):
            print(f"\n{'='*60}")
            print(f"Processing chart {i}/{len(self.unique_parents)}: {parent_name}")
            
            chart = RadarChart35(parent_name, self.df, self.text_processor)
            chart.create_chart()
            
            filename = f'radar_chart_35_{i}.png'
            chart.save_chart(self.output_folder, filename)
            print(f"✅ Saved: {filename}")
            
            with open(mapping_file, 'a', encoding='utf-8') as f:
                f.write(f"{filename} = {parent_name}\n")


# ==================== MAIN APPLICATION ====================

class AssessmentChartApp:
    """Main application orchestrating all chart types"""
    
    def __init__(self):
        self.file_path = None
        self.data_loader = None
        self.font_manager = FontManager()
        self.text_processor = PersianTextProcessor()
        
    def get_file_path(self):
        """Get file path from user"""
        user_input = input("Enter file path (with quotes): ")
        self.file_path = user_input.strip('"')
        return self.file_path
    
    def run_radar_31(self):
        """Generate radar charts from sheet 3.1"""
        print("\n" + "="*60)
        print("GENERATING RADAR CHARTS (SHEET 3.1)")
        print("="*60)
        
        self.data_loader = DataLoader(self.file_path)
        df, categories, values_matrix = self.data_loader.load_radar_data(sheet_name='3.1')
        
        self.font_manager.set_persian_font()
        
        chart = RadarChart31(df, categories, values_matrix)
        chart.draw()
        
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "radar_charts_31")
        chart.save_chart(output_folder, 'radar_chart_31.png')
    
        
        print("\n✅ Radar chart created successfully!")
    
    def run_bar_33(self):
        """Generate bar charts from sheet 3.3"""
        print("\n" + "="*60)
        print("GENERATING BAR CHARTS (SHEET 3.3)")
        print("="*60)
        
        self.data_loader = DataLoader(self.file_path)
        df = self.data_loader.load_bar_chart_data(sheet_name='3.3')
        
        config = {
            'chart_width': 6,
            'chart_height': 4,
            'bar_width': 0.9,
            'decimal_places': 1,
            'save_charts': True,
            'save_format': 'png',
            'dpi': 300
        }
        
        generator = BarChartGenerator33(df)
        generator.generate_all_charts(config)
    
    def run_assessment_34(self):
        """Generate assessment charts from sheet 3.4"""
        print("\n" + "="*60)
        print("GENERATING ASSESSMENT CHARTS (SHEET 3.4)")
        print("="*60)
        
        self.data_loader = DataLoader(self.file_path)
        df, categories, value_cols = self.data_loader.load_assessment_data(sheet_name='3.4')
        
        config = {
            'chart_width': 6,
            'chart_height': 4,
            'bar_width': 0.8,
            'decimal_places': 1,
            'multiply_by_100': True,
            'save_charts': True,
            'save_format': 'png',
            'dpi': 300
        }
        
        generator = AssessmentChartGenerator34(df, categories, value_cols)
        generator.generate_all_charts(config)
    
    def run_radar_35(self):
        """Generate radar charts from sheet 3.5"""
        print("\n" + "="*60)
        print("GENERATING RADAR CHARTS (SHEET 3.5)")
        print("="*60)
        
        generator = RadarChartGenerator35(self.file_path)
        generator.load_and_preprocess()
        generator.generate_all_charts()
    
    def run_all_without_prompt(self):
        """Run all chart generation methods without prompting for file path"""
        print("\n" + "="*60)
        print("ASSESSMENT CHART GENERATOR - UNIFIED")
        print("="*60)
        print("This will generate charts for sheets: 3.1, 3.3, 3.4, 3.5")
        print("="*60)
        
        try:
            self.run_radar_31()
        except Exception as e:
            print(f"⚠ Error in radar chart (3.1): {e}")
        
        try:
            self.run_bar_33()
        except Exception as e:
            print(f"⚠ Error in bar chart (3.3): {e}")
        
        try:
            self.run_assessment_34()
        except Exception as e:
            print(f"⚠ Error in assessment chart (3.4): {e}")
        
        try:
            self.run_radar_35()
        except Exception as e:
            print(f"⚠ Error in radar chart (3.5): {e}")
        
        print("\n" + "="*60)
        print("✅ All chart generation completed!")
        print("="*60)

# ==================== MAIN EXECUTION ====================

def main():
    """Main execution function"""
    app = AssessmentChartApp()
    
    print("="*60)
    print("UNIFIED ASSESSMENT CHART GENERATOR")
    print("="*60)
    print("This tool generates charts from Excel sheets 3.1, 3.3, 3.4, and 3.5")
    print("")
    print("Select an option:")
    print("1. Run all chart types")
    print("2. Run only Radar Chart (3.1)")
    print("3. Run only Bar Chart (3.3)")
    print("4. Run only Assessment Chart (3.4)")
    print("5. Run only Radar Chart (3.5)")
    print("="*60)
    
    choice = input("Enter your choice (1-5): ").strip()
    
    # Get file path ONCE before running any chart
    app.get_file_path()
    
    if choice == '1':
        # Run all but don't ask for file path again
        app.run_all_without_prompt()
    elif choice == '2':
        app.run_radar_31()
    elif choice == '3':
        app.run_bar_33()
    elif choice == '4':
        app.run_assessment_34()
    elif choice == '5':
        app.run_radar_35()
    else:
        print("Invalid choice. Running all charts...")
        app.run_all_without_prompt()

if __name__ == "__main__":
    main()