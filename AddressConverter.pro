QT       += core gui axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = AddressConverter
TEMPLATE = app

SOURCES += main.cpp\
        mainwindow.cpp \
    findbid.cpp \
    worker.cpp

HEADERS  += mainwindow.h \
    findbid.h \
    worker.h

FORMS    += mainwindow.ui \
    findbid.ui

DISTFILES += \
    README.md
