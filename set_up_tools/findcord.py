from pynput.mouse import Listener

#USED TO FIND COORDINATES OF SAP ENTRY
#RESULTS WILL BE USED IN THE PRINT_LINES METHOD
#AS HYPERPARAMETER

def on_click(x, y,button, pressed):
    if pressed:  # Only print when the mouse button is pressed
        print(f"Mouse clicked at X={x}, Y={y}")

# Start the mouse listener
with Listener(on_click=on_click) as listener:
    print("Click anywhere to get the coordinates. Press Ctrl+C to stop.")
    listener.join()
