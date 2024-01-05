# dmx-simples

### Description
C# implementation of the DMX-512 protocol (with a simple GUI 8 devices with 5 channel each)

### USB to RS485 Dongle
For this software to work, you must either USB to RS485 dongle (tested with XR21B142x). There are many on sell today, both based on the Prolific IC or FTDI. As the DMX-512 Protocol uses a high baud rate (250kb/s), it's recommended that you choose the FTDI one as the Prolific-based can't reach such high values.
