import matplotlib.pyplot as plt
import io
import re
import sys # To read multi-line input more robustly

def parse_data_block(data_text, num_columns):
    """Parses a block of text containing numerical data with a specific number of columns."""
    parsed_data = [[] for _ in range(num_columns)] # Create list of lists for columns
    lines = data_text.strip().split('\n')

    # Regex to find float-like numbers (including scientific notation)
    float_pattern = r'[\d\.\-eE]+'
    # Construct regex for the expected number of columns separated by whitespace
    regex_pattern = r"^\s*" + r"\s+".join([f'({float_pattern})'] * num_columns) + r"\s*$"
    compiled_regex = re.compile(regex_pattern)

    for line in lines:
        line = line.strip()
        if not line:
            continue
        match = compiled_regex.match(line)
        if match:
            try:
                for i in range(num_columns):
                    parsed_data[i].append(float(match.group(i+1)))
            except ValueError:
                print(f"Warning: Could not parse line as float data: {line}")
                continue
        else:
             # Add a check for lines that might just be indices or separators from partial pastes
             if not re.match(r"^\d+\s*$", line) and not "---" in line : # Ignore lines that are just integers (like indices) or separators
                print(f"Warning: Line format unexpected for {num_columns} columns: {line}")

    # Check if any data was actually parsed
    if not any(parsed_data):
        print(f"Warning: No valid data lines found for {num_columns} columns in the provided block.")
        return None # Indicate failure

    # Check if all columns have the same length (consistency)
    if len(parsed_data[0]) == 0:
         print(f"Warning: No valid data rows found for {num_columns} columns.")
         return None
    first_len = len(parsed_data[0])
    if not all(len(col) == first_len for col in parsed_data):
        print(f"Warning: Inconsistent number of rows parsed for {num_columns} columns.")
        # Decide how to handle this - perhaps return None or the shortest length data?
        # For now, let's try returning what we got, but it might cause issues later.
        min_len = min(len(col) for col in parsed_data)
        parsed_data = [col[:min_len] for col in parsed_data]
        if min_len == 0:
            return None


    return parsed_data

def get_multiline_input(prompt):
    """Helper function to get multi-line input from the user."""
    print(prompt)
    print("Paste your data. Press Enter twice or Ctrl+D (Unix) / Ctrl+Z+Enter (Windows) to finish.")
    lines = []
    while True:
        try:
            line = input()
            # Allow stopping with an empty line only if some lines were already entered
            if not line and lines:
                break
            elif line:
                lines.append(line)
        except EOFError: # Handle Ctrl+D/Ctrl+Z
            break
    print("--- End of input ---")
    return "\n".join(lines)

def plot_lagrange(original_x, original_y, evaluated_xp, evaluated_p):
    """Generates the plot for Lagrange interpolation."""
    plt.figure(figsize=(10, 6))
    plt.plot(original_x, original_y, 'ro', label='Original Data Points (yi)')
    plt.plot(evaluated_xp, evaluated_p, 'b.--', label='Lagrange Polynomial P(xp)')
    plt.title('Lagrange Interpolation')
    plt.xlabel('x')
    plt.ylabel('y / P(x)')
    plt.legend()
    plt.grid(True)
    # plt.show() # Show is called later

def plot_spline(nodes_xi, nodes_yi, mid_x, mid_s, mid_f, mid_error):
    """Generates the plots for Cubic Spline interpolation."""
    # --- Plot 1: Spline vs Exact Function ---
    plt.figure(figsize=(12, 7))
    plt.plot(nodes_xi, nodes_yi, 'ro', label='Interpolation Nodes (yi = f(xi))', markersize=8)
    # Corrected line below:
    plt.plot(mid_x, mid_s, 'g^', label='Spline S(x_mid)', markersize=6, linestyle='--')
    plt.plot(mid_x, mid_f, 'b.-', label='Exact f(x_mid)', markersize=6)
    plt.title('Cubic Spline Interpolation vs Exact Function')
    plt.xlabel('x')
    plt.ylabel('y / S(x) / f(x)')
    plt.legend()
    plt.grid(True)
    # Don't call plt.show() yet

    # --- Plot 2: Absolute Error ---
    plt.figure(figsize=(12, 5))
    plt.plot(mid_x, mid_error, 'm.-', label='Absolute Error |S(x_mid) - f(x_mid)|')
    plt.title('Cubic Spline Absolute Error at Midpoints')
    plt.xlabel('x_mid')
    plt.ylabel('Absolute Error |S - f|')
    # Check if error data is suitable for log scale
    # Use a small epsilon to avoid issues with exact zeros if log scale is desired
    epsilon = 1e-15
    if mid_error and all(e > epsilon for e in mid_error if e is not None):
         plt.yscale('log')
         plt.grid(True, which='both', linestyle='--')
    else:
         plt.grid(True)
         if any(e <= epsilon for e in mid_error if e is not None):
              print("Note: Log scale for error plot skipped due to zero or near-zero errors.")

    plt.tight_layout()

# --- Main Script Execution ---

print("="*40)
print(" Step 1: Lagrange Original Data")
print("="*40)
lagrange_original_text = get_multiline_input(
    "Paste ONLY the original (xi, yi) data points from Lagrange (2 columns):"
)
lagrange_orig_data = parse_data_block(lagrange_original_text, 2)
l_orig_x, l_orig_y = lagrange_orig_data if lagrange_orig_data else ([], [])
print(f"Parsed {len(l_orig_x)} Lagrange original points.")

print("\n" + "="*40)
print(" Step 2: Lagrange Evaluated Data")
print("="*40)
lagrange_evaluated_text = get_multiline_input(
    "Paste ONLY the evaluated (xp, P(xp)) data points from Lagrange (2 columns):"
)
lagrange_eval_data = parse_data_block(lagrange_evaluated_text, 2)
l_eval_x, l_eval_p = lagrange_eval_data if lagrange_eval_data else ([], [])
print(f"Parsed {len(l_eval_x)} Lagrange evaluated points.")


print("\n" + "="*40)
print(" Step 3: Spline Node Data")
print("="*40)
spline_nodes_text = get_multiline_input(
    "Paste ONLY the Spline Nodes (xi, yi, Mi) data (3 columns):"
)
spline_node_data = parse_data_block(spline_nodes_text, 3)
s_nodes_x, s_nodes_y, s_nodes_m = spline_node_data if spline_node_data else ([], [], [])
print(f"Parsed {len(s_nodes_x)} Spline nodes.")


print("\n" + "="*40)
print(" Step 4: Spline Comparison Data")
print("="*40)
spline_comparison_text = get_multiline_input(
    "Paste ONLY the Spline Comparison (x_mid, S(x_mid), f(x_mid), |S-f|) data (4 columns):"
)
spline_comp_data = parse_data_block(spline_comparison_text, 4)
s_mid_x, s_mid_s, s_mid_f, s_mid_err = spline_comp_data if spline_comp_data else ([], [], [], [])
print(f"Parsed {len(s_mid_x)} Spline comparison points.")


# --- Generate Plots ---
plots_generated = False
print("\n" + "="*40)
print("Generating Plots...")
print("="*40)

# Plot Lagrange if both original and evaluated data are present
if l_orig_x and l_eval_x:
    print("Generating Lagrange plot...")
    try:
        plot_lagrange(l_orig_x, l_orig_y, l_eval_x, l_eval_p)
        plots_generated = True
    except Exception as e:
        print(f"Error generating Lagrange plot: {e}")
else:
    print("Skipping Lagrange plot: Missing original or evaluated data.")
    if not l_orig_x: print(" -> Lagrange original data missing.")
    if not l_eval_x: print(" -> Lagrange evaluated data missing.")


# Plot Spline if both node and comparison data are present
if s_nodes_x and s_mid_x:
    print("Generating Spline plots...")
    try:
        plot_spline(s_nodes_x, s_nodes_y, s_mid_x, s_mid_s, s_mid_f, s_mid_err)
        plots_generated = True
    except Exception as e:
        print(f"Error generating Spline plots: {e}")
else:
    print("Skipping Spline plots: Missing node or comparison data.")
    if not s_nodes_x: print(" -> Spline node data missing.")
    if not s_mid_x: print(" -> Spline comparison data missing.")


# --- Show Plots ---
if plots_generated:
    print("\nDisplaying plots...")
    plt.show()
else:
    print("\nNo plots were generated.")

print("\nPlotting finished.")