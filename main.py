# main.py
import sys
import os

# Add current directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def main():
    """Main entry point untuk aplikasi Document Converter"""
    
    if len(sys.argv) > 1:
        # CLI Mode
        from cli.cli_converter import CLIConverter
        converter = CLIConverter()
        converter.run()
    else:
        # GUI Mode
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
            
            # Fallback ke CLI help
            if len(sys.argv) == 1:
                sys.argv.append('--help')
                from cli.cli_converter import CLIConverter
                converter = CLIConverter()
                converter.run()


if __name__ == "__main__":
    main()