# Java Access Bridge API UDF

## Table of Contents
+ [About](#about)
+ [Prerequisites](#prerequisites)
+ [Installation](#installation)
+ [Usage](#usage)
+ [Documentation](#documentation)


## About <a name = "about"></a>
This UDF is an AutoIt wrapper for 'JAB API'.   
Java Access Bridge (JAB) is a technology that exposes the Java Accessibility API in a Microsoft Windows DLL, enabling Java applications and applets that implement the Java Accessibility API to be visible to assistive technologies on Microsoft Windows systems.   
```Note, that not all functionality might be mapped or updated in current published version.```

## Prerequisites <a name = "prerequisites"></a>
Java Runtime Environment 'JRE 8/9 1.8.0' or higher and windowsaccessbridge.dll

## Installation <a name = "installation"></a>
Simply copy ```JavaAccessBridge_*_UDF.au3``` files to your development directory and use ```#include``` to point to these files in the source code.  

## Usage <a name = "usage"></a>
Use [JavaAccessBridgeExplorer](https://github.com/google/access-bridge-explorer) tool with JAB enabled to find windows and controls to develop automation using this UDF

## Documentation <a name = "documentation"></a>
* [JAB API Documentation](./docs/JSACC.pdf)
* [Installing Java Access Bridge](https://docs.oracle.com/javase/accessbridge/2.0.2/setup.htm#jab-installing-java-access-bridge])
* [Enable and test JAB](https://docs.oracle.com/javase/7/docs/technotes/guides/access/enable_and_test.html)
* [JAB introduction](https://docs.oracle.com/javase/accessbridge/2.0.2/introduction.htm)
* [Package summary](https://docs.oracle.com/javase/6/docs/api/javax/accessibility/package-summary.html)
* [JAB API](https://docs.oracle.com/javase/8/docs/technotes/guides/access/jab/api.html)
* [JS Accessibility](https://docs.oracle.com/javase/9/access/jaapi.htm#JSACC-GUID-0F6F322C-59E9-4190-AC13-BB78B52E2D5C)
* [Java Ferret](https://docs.oracle.com/javase/accessbridge/2.0.2/javaferret.htm)
* [AutoIt Java simple spy](https://www.autoitscript.com/forum/topic/166830-java-object-automation-and-simple-spy/page/3/)
* [AutoIt Java Windows Controls](https://www.autoitscript.com/forum/topic/42691-how-to-automate-java-windowscontrols/)
* [AutoIt Java based GUI](https://www.autoitscript.com/forum/topic/52853-scripting-to-a-java-based-gui/)
* [Stack Overflow JAB](https://stackoverflow.com/questions/19057985/automation-using-java-access-bridge)