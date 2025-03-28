from colorama import Fore, Style, init

init()

def show_banner():
    banner = """
  ██████╗  ██╗  ██╗ ██████╗ ███████╗████████╗██████╗ ███████╗██╗      █████╗ ██╗   ██╗
 ██╔════╝  ██║  ██║██╔═══██╗██╔════╝╚══██╔══╝██╔══██╗██╔════╝██║     ██╔══██╗╚██╗ ██╔╝
 ██║  ███╗███████║██║   ██║███████╗   ██║   ██████╔╝█████╗  ██║     ███████║ ╚████╔╝ 
 ██║   ██║██╔══██║██║   ██║╚════██║   ██║   ██╔══██╗██╔══╝  ██║     ██╔══██║  ╚██╔╝  
 ╚██████╔╝██║  ██║╚██████╔╝███████║   ██║   ██║  ██║███████╗███████╗██║  ██║   ██║   
  ╚═════╝ ╚═╝  ╚═╝ ╚═════╝ ╚══════╝   ╚═╝   ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝   ╚═╝   
    """
    colored_banner = ""
    for char in banner:
        if char != ' ' and char != '\n':
            colored_banner += Fore.GREEN + char
        else:
            colored_banner += char
    print(colored_banner + Style.RESET_ALL)

def show_status(message, status="info"):
    if status == "success":
        print(f"{Fore.GREEN}✓ {message}{Style.RESET_ALL}")
    elif status == "error":
        print(f"{Fore.RED}✗ {message}{Style.RESET_ALL}")
    elif status == "warning":
        print(f"{Fore.YELLOW}⚠ {message}{Style.RESET_ALL}")
    else:
        print(f"{Fore.CYAN}ℹ {message}{Style.RESET_ALL}")