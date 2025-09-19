from app_builder import AppBuilder

if __name__ == "__main__":
    """ 어플리케이션 시작 포인트 """
    builder = AppBuilder()
    builder.make_app()
    builder.exec_app()