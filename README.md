# EHKnjiga
EHKnjiga je program za generiranje englesko hrvatskog rječnika na temelju slobodno dostupnih rječničkih baza EH.Txt, HR.Txt i EN.Txt

Program je napravljen u Microsoft Accessu, ali na način da je sav programski kod pohranjen u tekstualnim datotekama kako bi se mogao prevesti u neki drugi programski jezik. Zbog toga nisu korištene mnoge mogućnosti Microsoft Accessa. Paralelno s izradom EH Knjige pišem i opsežni dokument u kojem je detaljno objašnjen program za izradu EH Knjige. Ovaj dokument je u nastajanju i mogu ga poslati na zahtjev

## Priprema EH Knjige
Najprije je potrebno U MS Access datoteku uvesti kopletan programski kod. Pokrene se (dvostrukim klikom) datoteka EHKnjiga.accdb 
Nakon toga se izvede naredba: Database Tools > Visual Basic

Nakon toga se u prozoru za zadavanje naredbi (Immediate Window) upiše naredba 

Call LoadFromText(acModule, "Admin", CurrentProject.Path & "\kod\Admin.txt")
