# EHKnjiga
EHKnjiga je program za generiranje englesko hrvatskog rječnika na temelju slobodno dostupnih rječničkih baza EH.Txt, [__HR.Txt__](https://github.com/gigaly/rjecnik-hrvatskih-jezika) i EN.Txt

Program je napravljen u Microsoft Accessu, ali na način da je sav programski kod pohranjen u tekstualnim datotekama kako bi se mogao prevesti u neki drugi programski jezik. Zbog toga nisu korištene mnoge mogućnosti Microsoft Accessa. Paralelno s izradom EH Knjige pišem i opsežni dokument u kojem je detaljno objašnjen program za izradu EH Knjige. Ovaj dokument je u nastajanju i mogu ga poslati na zahtjev

## Priprema EH Knjige
Najprije je potrebno U MS Access datoteku uvesti kompletan programski kod. Pokrene se (dvostrukim klikom) datoteka __EHKnjiga.accdb__ 
Nakon toga se izvede naredba: __Database Tools > Visual Basic__

Nakon toga se u prozoru za zadavanje naredbi (_Immediate Window_) upišu ove tri naredbe 

__Call LoadFromText(acModule, "Admin", CurrentProject.Path & "\kod\Admin.txt")__

__Call Admin.dodaj_reference__

__Call Admin.uvezi_sve(11)__

To je dosta za prvi puta. Možete analizrati kod, a ja nastavljam objašnjavati malo po malo


