        ��  ��                  ]   4   T E X T   A S S E M C L         0	                <comClass
            clsid="[CLSID]"
            threadingModel = "Apartment" />
   x   4   T E X T   A S S E M F F         0	                <typelib
            tlbid="[LIBID]"
            version="[VERSION]"
            helpdir="" />
    </file>
   4   T E X T   A S S E M F H         0	            <file name = "[LIBFILE]">
    4   T E X T   A S S E M F T         0	        </assembly>
     4   T E X T   A S S E M H D         0	        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <assemblyIdentity
        name="[ASSEMNAME]"
        processorArchitecture="X86" 
        type="win32"
        version="[VERSION]" />
 �  0   T E X T   A P P M A N       0	        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <assemblyIdentity 
        name="[APPNAME]" 
        processorArchitecture="X86" 
        type="win32"
        version="[VERSION]" />
    [APPDESC]
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                language="*"
                name="Microsoft.Windows.Common-Controls"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                type="win32"
                version="6.0.0.0" />
        </dependentAssembly>
    </dependency>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                name="[ASSEMNAME]"
                processorArchitecture="X86"
                type="win32"
                version="[VERSION]" />
        </dependentAssembly>
    </dependency>
</assembly>  