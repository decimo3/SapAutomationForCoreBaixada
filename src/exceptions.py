"""This module contains exceptions used in the SapWrapper library."""

class WrapperBaseException(Exception):
  ''' Base class for all exception on SAP_BOT '''
  def __init__(self, message: str) -> None:
    self.message = message
  def __str__(self) -> str:
    return self.message

class UnavailableSap(WrapperBaseException):
  ''' Exception class when we are unable to connect to SAP GUI Scripting Engine '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class SomethingGoesWrong(WrapperBaseException):
  ''' Exception class when something programactily goes wrong '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class ElementNotFound(WrapperBaseException):
  ''' Exception class when we are unable to find element '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class InformationNotFound(WrapperBaseException):
  ''' Exception class when requested information is missing '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class ArgumentException(WrapperBaseException):
  ''' Exception class when argument is invalid '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class TooMannyRequests(WrapperBaseException):
  ''' Exception class when have too many things to process '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))
  
class UnavailableTransaction(WrapperBaseException):
  ''' Exception class when we are unable to connect to SAP GUI Scripting Engine '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))
