class ApplicationError(Exception):
    """可安全返回给接口调用方的业务错误。"""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message
