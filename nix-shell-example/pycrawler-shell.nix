# packages repo you can find at
# https://nixos.org/nixos/pac5kages.html?channel=nixos-19.03
with import <nixpkgs> {};

let
    mypython = python37;
in
    stdenv.mkDerivation {
    name = "pycrawler-devenv";

    buildInputs = [
        # for printing some fancy hello
        figlet
        lolcat
        fd

        mypython
        mypython.pkgs.pylint
        mypython.pkgs.beautifulsoup4
        mypython.pkgs.requests
    ];

    shellHook = ''
        figlet "hello!" | lolcat --freq 0.5
        pwd

        findnix(){
            fd $1 /nix/store
        }
        export -f findnix
    '';
    }
