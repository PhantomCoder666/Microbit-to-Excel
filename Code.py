import asyncio
from bleak import BleakClient
from openpyxl import load_workbook

# Replace with actual Bluetooth address of your micro:bit
MICROBIT_ADDRESS = "XX:XX:XX:XX:XX:XX"  # ← update this!

# UART RX characteristic UUID for micro:bit (do not change unless needed)
UART_RX = "6e400003-b5a3-f393-e0a9-e50e24dcca9e"

# Mapping of scenario keys to Excel cells
# Cange if you wnat more or less scenarios
# You can also chane scenarios if you want
# If your dumb the scenarios are the "x", "y", "z" and so on
scenario_map = {
    "x": "D4",
    "y": "D5",
    "z": "D6",
    "temp": "E4",
    "light": "E5",
    "custom1": "F4",
    "custom2": "F5",
}

# Function to write a value to a specific cell in the Excel file
def write_to_excel(scenario, value):
    try:
        wb = load_workbook("Design.xlsx")  # Load the Excel workbook # Ensure this file exists and cange Design.xlsx to your file name
        sheet = wb["Home"]                 # Select the 'Home' sheet
        cell = scenario_map.get(scenario)  # Get the cell for the scenario

        if not cell:
            print(f"[!] Key '{scenario}' is not mapped to any cell.")
            return

        sheet[cell] = value                # Write the value to the cell
        wb.save("Design.xlsx")             # Save the workbook
        print(f"[✓] Wrote {value} to cell {cell} (key: '{scenario}')")
    except Exception as e:
        print(f"[X] Failed to write to Excel: {e}")

# Main async function to connect and listen for data
async def main():
    print("🔍 Connecting to micro:bit...")
    async with BleakClient(MICROBIT_ADDRESS) as client:
        print("✅ Connected. Listening for data...")

        # Handler for received data
        def handle_rx(_, data: bytearray):
            try:
                message = data.decode("utf-8").strip()
                print("🔹 Received:", message)

                # Expecting messages in the format 'key:value'
                if ":" in message:
                    scenario, value = message.split(":")
                    scenario = scenario.strip()
                    value = value.strip()
                    write_to_excel(scenario, value)
                else:
                    print("[!] Invalid message format. Expected 'key:value'")

            except Exception as e:
                print(f"[X] Error while handling data: {e}")

        # Start notifications from the micro:bit
        await client.start_notify(UART_RX, handle_rx)

        try:
            while True:
                await asyncio.sleep(1)  # Keep the program running
        except KeyboardInterrupt:
            print("\n🛑 Disconnected from micro:bit.")

# Entry point
if __name__ == "__main__":
    asyncio.run(main())

# Note: Ensure you have the required libraries installed:
# remember to install bleak and openpyxl using pip:
# pip install bleak openpyxl
## Also, ensure your micro:bit is set up to send data in the expected format.
# Make sure your micro:bit is sending data in the format 'key:value'
# And by the way Nah Id win 
# Courtsety of the one and only Froggie Coder and The Best Coders In BGS Tom,Ben and Nik
