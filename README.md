Orbitron2Omnirig is a small DDE Driver for [Orbitron](https://www.stoff.pl/) that can be used to feed frequency data into [OmniRig](https://www.dxatlas.com/omnirig/) to control you radio via CAT. Here is the idea:  
**Orbitron** =[(DDE)](https://en.wikipedia.org/wiki/Dynamic_Data_Exchange)=> **Orbitron2Omnirig** =[(COM)](https://en.wikipedia.org/wiki/Component_Object_Model)=> **OmniRig** =[(CAT)](https://en.wikipedia.org/wiki/Computer_aided_transceiver)=> **radio**  
  
Orbitron requires an .exe DDE driver, so this Python script needs to be compiled into a standalone executable file as per the instructions below.

Each radio supports a different set of CAT commands. You might need to modify the code/logic of `rig_init` and `rig_set_freq` according to your radio capabilities.  
For example my radio (ICOM 706 MKIIG) does not support ___setting frequency on VfoB___ directly, it only supports ___activating VfoB___ and ___setting the frequency on the active Vfo___ (and then activating VfoA). 

### Build the exe
```bash
pyinstaller --clean --onefile --name Orbitron2Omnirig main.py
```
Move/copy the generated `Orbitron2Omnirig.exe` in the `{Orbitron}` directory (where `Orbitron.exe` is)

### Configure Orbitron
Configure the file `{Orbitron}\Config\Setup.cfg`
```ini
[System]
DDEdriver=Orbitron2Omnirig
```

The following is needed only if `Orbitron2Omnirig.exe` is not placed in the `{Orbitron}` directory.
```ini
[Drivers]
Orbitron2Omnirig=C:\path\to\Orbitron2Omnirig.exe
```

### Changelog
- v0.1.0: First version
- v0.2.0: Update DL frequency as soon as needed
- v0.3.0: Better handling of frequency adjust: we adjust active Vfo (A or B), always
