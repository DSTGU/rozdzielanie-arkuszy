Aby skrypt działał prawidłowo wymagane jest stworzenie arkusza w sheetsach o nazwie "Admin".

Struktura tego arkusza wygląda następująco

![schema](https://raw.githubusercontent.com/DSTGU/rozdzielanie-arkuszy/refs/heads/master/schema.png)

Należy uzupełnić prawidłowymi danymi następujące pola:
(pola pod podpisem, nie na czerwono)

- Arkusz obdzwanianych: Nazwa arkusza z którego będą brane numery telefonów (lub inne dane udostępniane dalej obdzwaniającym)
- Arkusz obdzwaniających: Nazwa arkusza z którego będą brane maile do udostępniania arkusze, w którym będą umieszczane linki do stworzonych arkuszy oraz w których będzie gromadzony postęp
- Telefony: Nazwa kolumny której dane będą udostępniane obdzwaniającym
- Maile: Nazwa kolumny z adresami email ludzi, którym udostępniane będą dane
- Arkusze: Nazwa kolumny na linki do arkuszy

Należy również uzupełnić schemę, gdzie pierwszy rząd oznacza nazwę kolumny, a następne to możliwe odpowiedzi, przy czym ostatni rząd oznacza domyślnie zaznaczoną wartość, a pierwszy rząd oznacza pożądaną wartość (tą, której istnienie będzie podliczane)

Schema zostanie dodana również do arkusza obdzwaniających
