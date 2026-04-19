QT += core gui network xml

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

SOURCES += \
    main.cpp \
    mainwindow.cpp \
    simpledocparser.cpp \
    simpletranslator.cpp

HEADERS += \
    mainwindow.h \
    simpledocparser.h \
    simpletranslator.h

FORMS += \
    mainwindow.ui