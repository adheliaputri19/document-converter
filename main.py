import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def main():
    """Main entry point untuk aplikasi Document Converter"""
    
    if len(sys.argv) > 1:

        from cli.cli_converter import CLIConverter
        converter = CLIConverter()
        converter.run()
    else:
        try:
            from ui.gui_manager import GUIManager
            app = GUIManager()
            app.run()
        except ImportError as e:
            print(f"❌ Error: Tidak dapat menjalankan GUI mode")
            print(f"   Detail: {e}")
            print("\n🔧 Coba install dependencies:")
            print("   pip install tkinter")
            print("\n📟 Atau gunakan CLI mode:")
            print("   python main.py --help")
            
            if len(sys.argv) == 1:
                sys.argv.append('--help')
                from cli.cli_converter import CLIConverter
                converter = CLIConverter()
                converter.run()


if __name__ == "__main__":
    main()