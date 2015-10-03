Attribute VB_Name = "def"
Option Explicit

Public Enum ErrorEnum
    
    ' standard VBA errors
    InvalidProcedureCall = 5
    Overflow = 6
    OutOfMemory = 7
    SubscriptOutOfRange = 9
    FixedLockedArray = 10
    DivisionByZero = 11
    TypeMismatch = 13
    OutOfStackSpace = 28
    SubFuncNotDefined = 35
    Internal = 51
    BadFile = 52
    FileNotFound = 53
    BadFileMode = 54
    FileAlreadyOpen = 55
    FileAccessError = 75
    PathNotFound = 76
    ObjectVariableNotSet = 91
    InvalidUseOfNull = 94
    InvalidFileFormat = 321
    InvalidPropertyValue = 380
    InvalidPropertyArrayIndex = 381
    ReadOnlyProperty = 383
    WriteOnlyProperty = 394
    PropertyNotFound = 422
    PropertyOrMethodNotFound = 423
    ObjectRequired = 424
    UnsupportedPropertyOrMethod = 438
    Automation = 440
    ArgumentNotOptional = 449
    InvalidPropertyAssignment = 450
    CollectionKeyExists = 457
    UnsupportedEvents = 459
    MethodOrDataMemberNotFound = 461
    ApplicationOrObjectDefined = 1004
    
    ' custom errors
    BadFunctionCall = 35
    LogicError = 551
    RuntimeError = 591
    InvalidArgument = 593
    
End Enum
