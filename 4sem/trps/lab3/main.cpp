#include <iostream>
#include <string>

struct rec{
    std::string FIO;
    short Year;
    std::string Adr;
    std::string Tel;
};

int main(){
    FILE* Ftp;
    long Mark;
    rec V;
    std::string s, Fn;
    char c;

    std::cin >> Fn;
    Ftp = fopen(Fn.c_str(), "r");

    if (Ftp){
        Mark = 0;
        while (fread(&V, sizeof(V),1, Ftp)!=EOF){
            std::cout << V.FIO <<std::endl;
            std::cout << V.Year << std::endl;
            std::cout << V.Adr << std::endl;
            std::cout << V.Tel << std::endl;
        }
    }
}
