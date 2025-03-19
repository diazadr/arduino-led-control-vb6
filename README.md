# LED Control via Arduino & Visual Basic

![Project Status](https://img.shields.io/badge/status-completed-brightgreen) [![License](https://img.shields.io/badge/license-MIT-blue)](./LICENSE)

This tutorial explains how to control the ON/OFF state of an LED on Arduino pin 13 using buttons in Visual Basic.

## Technologies Used
- **Arduino Board**: The hardware used.
- **Arduino IDE**: Used to write and upload code to the Arduino.
- **Visual Basic**: Used to build the application interface.

## Features
- **LED Control**: Turn the LED on and off using buttons in Visual Basic.
- **Serial Communication**: Uses MSComm for communication between Visual Basic and Arduino.
- **Simple Interface**: Provides an easy-to-use interface for controlling the LED.

## Demo

### **Arduino Circuit**
Connect the LED to pin 13 on the Arduino and connect the Arduino to the PC.

#### **Arduino Code**  
The following code is used to receive commands from Visual Basic via serial communication:
```cpp
void setup() {
  Serial.begin(9600);
  pinMode(13, OUTPUT);
}

void loop() {
  if (Serial.available()) {
    char a = Serial.read();
    if (a == '1') {
      digitalWrite(13, HIGH);
    } else if (a == '0') {
      digitalWrite(13, LOW);
    }
    delay(100);
  }
}
```

### **MSComm Setup in Visual Basic**
To enable communication between Visual Basic and Arduino, configure MSComm with the following steps:
1. Click **Project**, select **Components** (or press **Ctrl + T**).
2. Check **Microsoft Comm Control 6.0 (SP6)** and click **OK**.
3. Add **two buttons** for turning the LED on and off, and **one MSComm** to the Visual Basic form.

### **Visual Basic Form Interface**  
Below is an example of the Visual Basic user interface:

<div align="center">
  <img src="https://github.com/user-attachments/assets/3477a17e-f892-4741-8e43-3213e667f368" alt="Form Interface Screenshot" width="600">
</div>

### **LED Control Result**  
Below is an example of the LED output when controlled via Visual Basic:

<div align="center">
  <img src="https://github.com/user-attachments/assets/5546adc9-5390-4529-91dc-65c828e0d11c" alt="LED Control Result" width="600">
</div>

<div align="center">
  <img src="https://github.com/user-attachments/assets/40c03c4a-89a6-40ca-994e-894bee1fbc17" alt="LED Control Result" width="600">
</div>

### **Port Configuration in Visual Basic**
Ensure that the communication port between Arduino and Visual Basic matches. To find the Arduino port:
- **Arduino IDE**: Click **Tools** → **Port**.
- **Device Manager**: Check under **Ports (COM & LPT)**.

Set **CommPort** in MSComm to match the port used by Arduino.

### **Serial Configuration in Visual Basic**
Ensure that the serial communication configuration in Visual Basic matches the Arduino settings:
- Baudrate = 9600
- Data = 8
- Parity = None
- Stop bit = 1

## Installation and Usage
1. **Connect the Arduino to the PC** and upload the Arduino code.
2. **Set up MSComm in Visual Basic** for serial communication.
3. **Run the Visual Basic application** and control the LED using the available buttons.

## Project Status
This project is **completed** and will not be further developed.

## Contributions
Feel free to submit an **issue** or create a **pull request** if you wish to contribute.

## License
This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for more details.
