"""This module contains exceptions used in the SapWrapper library."""

class UnavailableSap(Exception):
  ''' Exception class when we are unable to connect to SAP GUI Scripting Engine '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class SomethingGoesWrong(Exception):
  ''' Exception class when something programactily goes wrong '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class ElementNotFound(Exception):
  ''' Exception class when we are unable to find element '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class InformationNotFound(Exception):
  ''' Exception class when requested information is missing '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class ArgumentException(Exception):
  ''' Exception class when argument is invalid '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))

class TooMannyRequests(Exception):
  ''' Exception class when have too many things to process '''
  def __init__(self, message: str = "") -> None:
    super().__init__(message)
    self.message = message
  def __str__(self) -> str:
    return self.message.format(**vars(self))