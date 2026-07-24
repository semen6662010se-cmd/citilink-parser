from parser_core import main_parser

if __name__ == "__main__":
    try:
        main_parser()
        print("\nПроцесс парсинга завершен без системных ошибок.")
    except Exception as err:
        error_type = type(err).__name__
        print(f"\nКритическая ошибка исполнения: {error_type}")
        print(f"Системный вывод: {err}")
        print("Требуется отладка исходного кода.")
    finally:
        input("\nНажмите Enter для закрытия терминала...")