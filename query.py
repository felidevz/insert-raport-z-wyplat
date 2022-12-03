from decimal import Decimal


def as_curr(number: Decimal | int = 0) -> str:
    """Return number as currency string with two decimal places, spaces and commas"""
    return f'{number:,.2f}'.replace(',', ' ').replace('.', ',')


class Query:
    def __init__(self, connection, month, year):
        self.connection = connection
        self.cursor = self.connection.cursor()
        self.month = month
        self.year = year
        self.rows = []

    def first_query(self):
        query = f'''
            SELECT dzi_Nazwa, dzi_Analityka, IsNull(SUM(wyp_BruttoDuze - IsNull(B.SumaZus, 0) ), 0) AS [Wartosc]
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            INNER JOIN pr_Pracownik on pr_Id = up_IdPracownika
            INNER JOIN plb_UmowaDzialStanowisko ON udz_IdUmowy = up_Id AND 
            (lp_DataWyplaty BETWEEN udz_DataOd AND ISNULL(udz_DataDo, '29991231') )
            RIGHT JOIN sl_Dzial DZ on DZ.dzi_Id = udz_IdDzialu
            LEFT JOIN twsf_DzialAnalityka A ON A.dzi_Id = DZ.dzi_Id
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus
                FROM plb_WyplataSkladnik
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika
                WHERE sp_PlatnyPrzez = 1
                GROUP BY wyps_IdWyplaty
            ) B ON B.wyps_IdWyplaty = wyp_Id WHERE 
            (lp_Id IS NULL OR (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year}) )
            
            GROUP BY dzi_Nazwa, dzi_Analityka
            ORDER BY dzi_Nazwa
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        obcy_zlecenia = (0, 0, 0)

        for row in rows:
            self.rows.append(
                (
                    f'431-{row[1]}', as_curr(row[2]), as_curr(), row[0], 'Brutto duże', ''
                )
            )

            if row[0] == 'Obcy zlecenia':
                obcy_zlecenia = row
                if obcy_zlecenia[0] is None:
                    obcy_zlecenia = (0, obcy_zlecenia[1], obcy_zlecenia[2])
                if obcy_zlecenia[1] is None:
                    obcy_zlecenia = (obcy_zlecenia[0], 0, obcy_zlecenia[2])
                if obcy_zlecenia[2] is None:
                    obcy_zlecenia = (obcy_zlecenia[0], obcy_zlecenia[1], 0)
        return obcy_zlecenia

    def second_query(self, obcy_zlecenia):
        query = f'''
            SELECT ISNULL(SUM(ru_BruttoDuze-ru_Akord), 0) AS BruttoMale, ISNULL(SUM(ru_Akord), 0) AS ru_Akord,
            ISNULL(SUM(ru_BruttoDuze), 0) AS ru_BruttoDuze
            FROM plb_RachunekDoUmowyCP
            INNER JOIN sl_Kategoria ON kat_Id = ru_IdKategorii
            WHERE kat_Nazwa = 'RADA' AND
            (MONTH(ru_DataWystawienia) = {self.month} and YEAR(ru_DataWystawienia) = {self.year})
        '''
        self.cursor.execute(query)
        row1 = self.cursor.fetchone()

        self.rows.append(
            (
                f'431-041-000', as_curr(row1[0]), as_curr(), 'Rada Nadzorcza',
                'Wartość brutto duże z umów o dzieło (Rady) minus kwota z umowy o dzieło z pola rozliczenie akordów', ''
            )
        )

        query = f'''
            SELECT ISNULL(SUM(ru_BruttoDuze-ru_Akord), 0) AS BruttoMale, ISNULL(SUM(ru_Akord), 0) AS ru_Akord,
            ISNULL(SUM(ru_BruttoDuze), 0) AS ru_BruttoDuze
            FROM plb_RachunekDoUmowyCP
            LEFT JOIN sl_Kategoria on kat_Id = ru_IdKategorii
            WHERE (kat_Nazwa IS NULL OR kat_Nazwa <> 'RADA') AND 
            (MONTH(ru_DataWystawienia) = {self.month} and YEAR(ru_DataWystawienia) = {self.year})
        '''
        self.cursor.execute(query)
        row2 = self.cursor.fetchone()

        self.rows.append(
            (
                f'431-041-000', as_curr(row2[0]), as_curr(), 'Ryczałt cz', '', ''
            )
        )

        self.rows.append(
            (
                f'232-004', as_curr(), as_curr(row2[0]), 'Ryczałt cz', '', ''
            )
        )

        self.rows.append(
            (
                f'232-005', as_curr(), as_curr(row1[0]), 'Rada Nadzorcza',
                'Wartość brutto duże z umów o dzieło (Rady)', ''
            )
        )

        self.rows.append(
            (
                f'232-012', as_curr(), as_curr(obcy_zlecenia[2]), 'Fun. ubezp. obcy zlecenia', '', ''
            )
        )

    def third_query(self, obcy_zlecenia):
        query = f'''
            SELECT SUM(ISNULL(wyps_WartoscFin, 0) ) AS WynChor,
            SUM(wyp_BruttoDuze - ISNULL(A.SumaZus, 0) - ISNULL(wyps_WartoscFin, 0) ) AS Brutto
            FROM plb_Wyplata
            LEFT JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika  = 25
            INNER JOIN plb_ListaPlac on lp_Id  = wyp_IdListyPlac
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus
                FROM plb_WyplataSkladnik
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika
                WHERE sp_PlatnyPrzez = 1
                GROUP BY wyps_IdWyplaty
            ) A ON A.wyps_IdWyplaty = wyp_Id WHERE (MONTH(lp_Miesiac) = {self.month} and YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        row = self.cursor.fetchone()
        if row[0] is None:
            row = (0, row[1])
        if row[1] is None:
            row = (row[0], 0)

        self.rows.append(
            (
                '232-001', as_curr(), as_curr(row[1] - obcy_zlecenia[2]), 'Osob. fund. pł. z ZUS',
                'Osob. fund. pł. z ZUS - obcy zlecenia - inne odprawy ekon./emer. '
                'Wartość brutto duże bez wynagrodzenia chorobowego', ''
            )
        )

        self.rows.append(
            (
                '232-002', as_curr(), as_curr(row[0]), 'Wynagrodzenie chorobowe', 'Wynagrodzenie chorobowe', ''
            )
        )

    def fourth_query(self):
        query = f'''
            SELECT ISNULL(SUM(wyp_BruttoDuze - ISNULL(A.SumaZus, 0) ), 0)
            FROM plb_Wyplata
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus FROM plb_WyplataSkladnik
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika
                WHERE sp_PlatnyPrzez = 1
                GROUP BY wyps_IdWyplaty
            ) A ON A.wyps_IdWyplaty = wyp_Id
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} and YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        row1 = self.cursor.fetchone()

        query = f'''
                SELECT ISNULL(SUM(ru_BruttoDuze), 0)
                FROM plb_RachunekDoUmowyCP
                INNER JOIN sl_Kategoria on kat_Id = ru_IdKategorii
                WHERE (MONTH(ru_DataWystawienia) = {self.month} AND YEAR(ru_DataWystawienia) = {self.year})
                '''
        self.cursor.execute(query)
        row2 = self.cursor.fetchone()

        self.rows.append(
            (
                '232-000', as_curr(row1[0] + row2[0]), as_curr(), 'Razem naliczenia',
                'Wartość brutto duże umowy o pracę i rachunków', ''
            )
        )

        self.rows.append(
            (
                '231-000', as_curr(), as_curr(row1[0] + row2[0]), 'Razem naliczenia',
                'Wartość brutto duże umowy o pracę i rachunków', ''
            )
        )

    def fifth_query(self):
        query = f'''
            SELECT dzi_Nazwa, dzi_Analityka,
            ISNULL(SUM(wyp_ZusEmer1Firma + wyp_ZusRentFirma + wyp_ZusWypFirma), 0) AS [Wartosc]
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            INNER JOIN pr_Pracownik ON pr_Id = up_IdPracownika
            INNER JOIN plb_UmowaDzialStanowisko ON udz_IdUmowy = up_Id AND
            (lp_DataWyplaty BETWEEN udz_DataOd AND ISNULL(udz_DataDo, '29991231') )
            RIGHT JOIN sl_Dzial DZ ON DZ.dzi_Id = udz_IdDzialu
            LEFT JOIN twsf_DzialAnalityka A ON A.dzi_Id = DZ.dzi_Id
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus
                FROM plb_WyplataSkladnik
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika
                WHERE sp_PlatnyPrzez = 1
                GROUP BY wyps_IdWyplaty
            ) B ON B.wyps_IdWyplaty = wyp_Id WHERE
            (lp_Id IS NULL OR (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year}) )
            
            GROUP BY dzi_Nazwa, dzi_Analityka
            ORDER BY dzi_Nazwa
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        for row in rows:
            self.rows.append(
                (
                    f'445-{row[1]}', as_curr(row[2]), as_curr(), row[0], 'Składki ZUS pracodawcy', ''
                )
            )

        query = f'''
            SELECT SUM(RUCP.ZUS_PRACOWNIK), SUM(RUCP.ZUS_PRACODAWCA)
            FROM vwGratRachunkiDoUmowCP AS RUCP
            LEFT JOIN fl_Wartosc AS FlagiWartosci ON RUCP.IDENT=FlagiWartosci.flw_IdObiektu AND flw_IdGrupyFlag = 3
            LEFT JOIN pd_Uzytkownik AS FlagaUzytk ON FlagiWartosci.flw_IdUzytkownika=FlagaUzytk.uz_Id
            LEFT JOIN fl__Flagi AS Flagi ON FlagiWartosci.flw_IdFlagi=Flagi.flg_Id
            WHERE (MONTH(DATA) = {self.month} AND YEAR(DATA) = {self.year}) AND (ID_KATEGORIA = 12)
        '''
        self.cursor.execute(query)
        zus_pracownik_pracodawca = self.cursor.fetchone()
        if zus_pracownik_pracodawca[0] is None:
            zus_pracownik_pracodawca = (0, zus_pracownik_pracodawca[1])
        if zus_pracownik_pracodawca[1] is None:
            zus_pracownik_pracodawca = (zus_pracownik_pracodawca[0], 0)

        query = f'''
            SELECT SUM(ru_FP)
            FROM plb_RachunekDoUmowyCP
            LEFT JOIN plb_UmowaCP ON ru_IdUmowy = ucp_Id
            LEFT JOIN pr_Pracownik ON ucp_IdPracownika = pr_Id
            WHERE (MONTH(ru_DataWystawienia) = {self.month} AND YEAR(ru_DataWystawienia) = {self.year})
        '''
        self.cursor.execute(query)
        fp_kwota = self.cursor.fetchone()[0]
        if fp_kwota is None:
            fp_kwota = 0

        self.rows.append(
            (
                '445-041-000', as_curr(zus_pracownik_pracodawca[1] - fp_kwota), as_curr(),
                'Rada ZUS pracodawcy bez FP', 'Składki ZUS pracodawcy', ''
            )
        )

        return (zus_pracownik_pracodawca[0], zus_pracownik_pracodawca[1], fp_kwota)

    def sixth_query(self, zus_pracownik_pracodawca_kwota_fp):
        query = f'''
            SELECT dzi_Nazwa, dzi_Analityka, ISNULL(SUM(wyp_FP), 0) AS [Wartosc]
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            INNER JOIN pr_Pracownik ON pr_Id = up_IdPracownika
            INNER JOIN plb_UmowaDzialStanowisko ON udz_IdUmowy = up_Id AND
            (lp_DataWyplaty BETWEEN udz_DataOd AND ISNULL(udz_DataDo, '29991231') )
            RIGHT JOIN sl_Dzial DZ ON DZ.dzi_Id = udz_IdDzialu
            LEFT JOIN twsf_DzialAnalityka A ON A.dzi_Id = DZ.dzi_Id
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus 
                FROM plb_WyplataSkladnik 
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika 
                WHERE sp_PlatnyPrzez = 1 
                GROUP BY wyps_IdWyplaty
            ) B ON B.wyps_IdWyplaty = wyp_Id WHERE
            (lp_Id IS NULL OR (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year}) )
            
            GROUP BY dzi_Nazwa, dzi_Analityka
            ORDER BY dzi_Nazwa
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        for row in rows:
            self.rows.append(
                (
                    f'445-{row[1]}', as_curr(row[2]), as_curr(), row[0], 'FP', ''
                )
            )

        self.rows.append(
            (
                '445-041-000', as_curr(zus_pracownik_pracodawca_kwota_fp[2]), as_curr(), 'Rada FP', 'FP', ''
            )
        )

    def seventh_query(self, zus_pracownik_pracodawca_kwota_fp):
        query = f'''
            SELECT ISNULL(SUM(wyp_ZusEmer1Firma + wyp_ZusRentFirma + wyp_ZusWypFirma), 0) AS ZusFirma,
            ISNULL(SUM(wyp_FP), 0) AS FP, ISNULL(SUM(wyp_FGSP), 0) AS FGSP
            FROM plb_Wyplata
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        row = self.cursor.fetchone()

        self.rows.append(
            (
                '229-051', as_curr(), as_curr(row[0]), 'Składka na ubezp. społ.',
                'Emerytalne, rentowe i wypadkowe płacone przez pracodawcę RAZEM', ''
            )
        )

        self.rows.append(
            (
                '229-051', as_curr(), as_curr(zus_pracownik_pracodawca_kwota_fp[1] - zus_pracownik_pracodawca_kwota_fp[2]),
                'Składka na ubezp. społ. RADA - FP', 'Emerytalne, rentowe płacone przez pracodawcę RAZEM', ''
            )
        )

        self.rows.append(
            (
                '229-053', as_curr(), as_curr(row[1]), 'Składka na fundusz pracy', 'Fundusz pracy RAZEM', ''
            )
        )

        self.rows.append(
            (
                '229-053', as_curr(), as_curr(zus_pracownik_pracodawca_kwota_fp[2]), 'Składka na fundusz pracy RADA',
                'Fundusz pracy RAZEM', ''
            )
        )

        self.rows.append(
            (
                '229-054', as_curr(), as_curr(row[2]), 'Składka FGŚP', 'Składka FGŚP RAZEM', ''
            )
        )

    def eighth_query(self):
        query = f'''
            SELECT dzi_Nazwa, dzi_Analityka, ISNULL(SUM(wyp_FGSP), 0) AS [Wartosc]
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            INNER JOIN pr_Pracownik ON pr_Id = up_IdPracownika
            INNER JOIN plb_UmowaDzialStanowisko ON udz_IdUmowy = up_Id AND
            (lp_DataWyplaty BETWEEN udz_DataOd AND ISNULL(udz_DataDo, '29991231') )
            RIGHT JOIN sl_Dzial DZ ON DZ.dzi_Id = udz_IdDzialu
            LEFT JOIN twsf_DzialAnalityka A ON A.dzi_Id = DZ.dzi_Id
            LEFT JOIN (
                SELECT wyps_IdWyplaty, SUM(wyps_WartoscFin) AS SumaZus
                FROM plb_WyplataSkladnik
                INNER JOIN plb_Skladnik ON sp_Id = wyps_IdSkladnika
                WHERE sp_PlatnyPrzez = 1
                GROUP BY wyps_IdWyplaty
            ) B ON B.wyps_IdWyplaty = wyp_Id WHERE
            (lp_Id IS NULL OR (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year}) )
            GROUP BY dzi_Nazwa, dzi_Analityka
            ORDER BY dzi_Nazwa
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        for row in rows:
            self.rows.append(
                (
                    f'445-{row[1]}', as_curr(row[2]), as_curr(), row[0], 'FGŚP', ''
                )
            )

    def ninth_query(self):
        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = 26
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        row1 = self.cursor.fetchone()

        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = 28
            INNER JOIN plb_ListaPlac ON lp_Id  = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        row2 = self.cursor.fetchone()

        query = f'''
            SELECT SUM(CASE WHEN SKL.sp_Rodzaj = 1 THEN - WYPS.wyps_WartoscFin ELSE WYPS.wyps_WartoscFin END) AS WARTOSC
            FROM plb_WyplataSkladnik WYPS
            LEFT OUTER JOIN dbo.plb_Skladnik AS SKL ON SKL.sp_Id = WYPS.wyps_IdSkladnika
            INNER JOIN plb_Wyplata WYP ON WYPS.wyps_IdWyplaty = WYP.wyp_Id
            INNER JOIN plb_ListaPlac LP ON WYP.wyp_IdListyPlac = LP.lp_Id
            INNER JOIN plb_Umowa UP ON UP.up_Id = WYP.wyp_IdUmowy
            INNER JOIN pr_Pracownik PR ON PR.pr_Id = UP.up_IdPracownika
            WHERE (MONTH(LP.lp_Miesiac) = {self.month} AND YEAR(LP.lp_Miesiac) = {self.year})
            AND LP.lp_Id = 580 AND WYPS.wyps_IdSkladnika = 27
        '''
        self.cursor.execute(query)
        row3 = self.cursor.fetchone()
        if row3[0] is None:
            row3 = (0, )

        self.rows.append(
            (
                '229-051', as_curr(row1[0]), as_curr(), 'Zasiłek ZUS 100%', 'Zasiłek chorobowy', ''
            )
        )

        self.rows.append(
            (
                '229-051', as_curr(row2[0]), as_curr(), 'Urlop macierzyński', 'Zasiłek macierzyński', ''
            )
        )

        self.rows.append(
            (
                '229-051', as_curr(row3[0]), as_curr(), 'Zasiłek opiekuńczy', 'Zasiłek opiekuńczy', ''
            )
        )

        self.rows.append(
            (
                '231-000', as_curr(), as_curr(row1[0] + row2[0] + row3[0]), 'Razem zasiłki ZUS',
                'Suma zasiłku chorobowego, macierzyńskiego i opiekuńczego', ''
            )
        )

    def tenth_query(self):
        query = f'''
            SELECT sp_Id
            FROM plb_Skladnik
            WHERE sp_Nazwa = 'Fundusz mieszkaniowy'
        '''
        self.cursor.execute(query)
        funduszm_id = self.cursor.fetchone()[0]
        return funduszm_id

    def eleventh_query(self, funduszm_id):
        query = f'''
            SELECT pr_WWW, pr_Imie + ' ' + pr_Nazwisko AS Pracownik, wyps_WartoscFin
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {funduszm_id}
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            INNER JOIN pr_Pracownik ON pr_Id = up_IdPracownika
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        for row in rows:
            self.rows.append(
                (
                    f'239-{row[0]}', as_curr(), as_curr(row[2]), row[1], 'Fundusz mieszkaniowy', row[0]
                )
            )

        query = f'''
            SELECT
            firmaPPK.Analityka,
            SUM(firmaPPK.PPK_pracodawcaPodst) as [PPK Firma podst],
            firmaPPK.Dzial
            FROM (
                SELECT 
                A.dzi_Analityka AS [Analityka],
                dzial1.dzi_Nazwa AS [Dzial],
                wypl1.[wyp_WplataPodstPracodawcyPPK] AS [PPK_pracodawcaPodst],
                wypl1.[wyp_WplataDodatkPracodawcyPPK] AS [PPK_pracodawcaDodatk]
                FROM [plb_Wyplata] AS wypl1
                LEFT JOIN plb_ListaPlac AS lista1 ON lista1.lp_Id = wypl1.wyp_IdListyPlac
                LEFT JOIN plb_UmowaDzialStanowisko AS umowadzialstan1 ON umowadzialstan1.udz_IdUmowy = wypl1.wyp_IdUmowy
                AND udz_DataOd <= lista1.lp_DataWyplaty AND (udz_DataDo IS NULL OR udz_DataDo >= lista1.lp_DataWyplaty)
                LEFT JOIN sl_Dzial AS dzial1 ON dzial1.dzi_Id = umowadzialstan1.udz_IdDzialu
                LEFT JOIN twsf_DzialAnalityka A ON A.dzi_Id = dzial1.dzi_Id
                WHERE (MONTH(lista1.lp_DataWyplaty) = {self.month} AND YEAR(lista1.lp_DataWyplaty) = {self.year})
            ) AS firmaPPK
            GROUP BY firmaPPK.Dzial, firmaPPK.Analityka
            ORDER BY firmaPPK.Dzial
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        ppk_pracodawca_razem = Decimal('0')

        for row in rows:
            ppk_pracodawca_razem += row[1]
            self.rows.append(
                (
                    f'413-{row[0][:3]}-005', as_curr(row[1]), as_curr(), f'PPK Pracodawca {row[2]}', '', ''
                )
            )

        self.rows.append(
            (
                '246-001', as_curr(), as_curr(ppk_pracodawca_razem), 'PPK Pracodawca RAZEM', '', ''
            )
        )

        query = f'''
            SELECT
            SUM(pracownikPPK.[PPK_pracownikPodst])
            FROM (
            SELECT 
            dzial1.dzi_Nazwa AS [Dzial],
            wypl1.[wyp_WplataPodstPracownikaPPK] AS [PPK_pracownikPodst],
            wypl1.[wyp_WplataDodatkPracownikaPPK] AS [PPK_pracownikDodatk]
            FROM [plb_Wyplata] AS wypl1
            LEFT JOIN plb_ListaPlac AS lista1 ON lista1.lp_Id = wypl1.wyp_IdListyPlac
            LEFT JOIN plb_UmowaDzialStanowisko AS umowadzialstan1 ON umowadzialstan1.udz_IdUmowy = wypl1.wyp_IdUmowy
            AND udz_DataOd <= lista1.lp_DataWyplaty AND (udz_DataDo IS NULL OR udz_DataDo >= lista1.lp_DataWyplaty)
            LEFT JOIN sl_Dzial AS dzial1 ON dzial1.dzi_Id = umowadzialstan1.udz_IdDzialu
            WHERE (MONTH(lista1.lp_DataWyplaty) = {self.month} AND YEAR(lista1.lp_DataWyplaty) = {self.year})
            ) AS pracownikPPK
            GROUP BY pracownikPPK.Dzial
            ORDER BY pracownikPPK.Dzial
        '''
        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        ppk_pracownik_razem = Decimal('0')

        for row in rows:
            ppk_pracownik_razem += row[0]

        self.rows.append(
            (
                '231-000', as_curr(ppk_pracownik_razem), as_curr(), 'PPK Pracownik RAZEM', '', ''
            )
        )

        self.rows.append(
            (
                '246-000', as_curr(), as_curr(ppk_pracownik_razem), 'PPK Pracownik RAZEM', '', ''
            )
        )

    def twelfth_query(self, funduszm_id):
        query = f'''
            SELECT ISNULL(SUM(wyps_WartoscFin), 0) AS FunduszMieszkRazem
            FROM plb_Wyplata
            INNER JOIN plb_Umowa ON up_Id = wyp_IdUmowy
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {funduszm_id}
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            INNER JOIN pr_Pracownik ON pr_Id = up_IdPracownika
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        funduszm_razem = self.cursor.fetchone()[0]
        return funduszm_razem

    def thirteenth_query(self):
        query = f'''
            SELECT ISNULL(SUM(ru_PotraceniaNetto), 0)
            FROM plb_RachunekDoUmowyCP
            WHERE (MONTH(ru_DataWystawienia) = {self.month} AND YEAR(ru_DataWystawienia) = {self.year})
        '''
        self.cursor.execute(query)
        potracenia = self.cursor.fetchone()[0]
        return potracenia

    def fourteenth_query(self):
        query = f'''
            SELECT sp_Id
            FROM plb_Skladnik
            WHERE sp_Nazwa = 'Składka na PZU'
        '''
        self.cursor.execute(query)
        pzu_id = self.cursor.fetchone()[0]
        return pzu_id

    def fifteenth_query(self, pzu_id):
        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {pzu_id}
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        pzu_razem = self.cursor.fetchone()[0]

        self.rows.append(
            (
                '249-006', as_curr(), as_curr(pzu_razem), 'PZU', 'Składka na PZU RAZEM', ''
            )
        )
        return pzu_razem

    def sixteenth_query(self):
        query = f'''
            SELECT sp_Id
            FROM plb_Skladnik
            WHERE sp_Nazwa = 'Ubezpieczenia szpitalne'
        '''
        self.cursor.execute(query)
        ubezp_szpitalne_id = self.cursor.fetchone()[0]
        return ubezp_szpitalne_id

    def seventeenth_query(self, ubezp_szpitalne_id):
        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {ubezp_szpitalne_id}
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        ubezp_szpitalne_kwota = self.cursor.fetchone()[0]

        self.rows.append(
            (
                '249-006', as_curr(), as_curr(ubezp_szpitalne_kwota), 'Ubezpieczenia szpitalne',
                'Ubezpieczenia szpitalne', ''
            )
        )
        return ubezp_szpitalne_kwota

    def eighteenth_query(self):
        query = f'''
            SELECT sp_Id
            FROM plb_Skladnik
            WHERE sp_Nazwa = 'Spłata pożyczki KZP'
        '''
        self.cursor.execute(query)
        splata_kzp_id = self.cursor.fetchone()[0]
        return splata_kzp_id

    def nineteenth_query(self, splata_kzp_id):
        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {splata_kzp_id}
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        splata_kzp_kwota = self.cursor.fetchone()[0]
        return splata_kzp_kwota

    def twentyth_query(self):
        query = f'''
            SELECT sp_Id
            FROM plb_Skladnik
            WHERE sp_Nazwa = 'Wkłady KZP'
        '''
        self.cursor.execute(query)
        wklady_kzp_id = self.cursor.fetchone()[0]
        return wklady_kzp_id

    def twentyfirst_query(self, wklady_kzp_id):
        query = f'''
            SELECT ISNULL(SUM(ISNULL(wyps_WartoscFin, 0) ), 0) AS Wartosc
            FROM plb_Wyplata
            INNER JOIN plb_WyplataSkladnik ON wyps_IdWyplaty = wyp_Id AND wyps_IdSkladnika = {wklady_kzp_id}
            INNER JOIN plb_ListaPlac on lp_Id = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        wklady_kzp_kwota = self.cursor.fetchone()[0]
        return wklady_kzp_kwota

    def twentysecond_query(self, splata_kzp_kwota, wklady_kzp_kwota):
        self.rows.append(
            (
                '249-003', as_curr(), as_curr(splata_kzp_kwota + wklady_kzp_kwota), 'PKZP',
                'Suma składników związanych z KZP', ''
            )
        )

    def twentythird_query(self):
        query = f'''
            SELECT ISNULL(SUM(wyp_ZaliczkaNaPodatek), 0) AS Podatek,
            ISNULL(SUM(wyp_Przelew), 0) AS Przelew, ISNULL(SUM(wyp_ZUSPrac), 0) AS ZusPrac,
            ISNULL(SUM(wyp_UbezpZdrowotne), 0) AS Zdrowotne7,
            ISNULL(SUM(wyp_UbezpZdrowotne + wyp_UbezpZdrowotne2), 0) AS Zdrowotne9
            FROM plb_Wyplata
            INNER JOIN plb_ListaPlac ON lp_Id = wyp_IdListyPlac
            WHERE (MONTH(lp_Miesiac) = {self.month} AND YEAR(lp_Miesiac) = {self.year})
        '''
        self.cursor.execute(query)
        podatki1 = self.cursor.fetchone()
        return podatki1

    def twentyfourth_query(self):
        query = f'''
            SELECT ISNULL(SUM(ru_ZaliczkaNaPodatek), 0) AS Podatek, ISNULL(SUM(ru_Przelew), 0) AS Przelew,
            ISNULL(SUM(ru_UbezpZdrowotne), 0) AS Zdrowotne7,
            ISNULL(SUM(ru_UbezpZdrowotne + ru_UbezpZdrowotne2), 0) AS Zdrowotne9
            FROM plb_RachunekDoUmowyCP
            WHERE (MONTH(ru_DataWystawienia) = {self.month} AND YEAR(ru_DataWystawienia) = {self.year})
        '''
        self.cursor.execute(query)
        podatki2 = self.cursor.fetchone()
        return podatki2

    def twentyfifth_query(self, podatki1, podatki2, potracenia_razem, zus_pracownik_pracodawca_kwota_fp):
        self.rows.append(
            (
                '220-002', as_curr(), as_curr(podatki1[0] + podatki2[0]), 'Podatek dochodowy', 'Podatek dochodowy', ''
            )
        )

        self.rows.append(
            (
                '249-004', as_curr(), as_curr(podatki1[1] + podatki2[1]), 'Konto bankowe', 'Konto bankowe', ''
            )
        )

        query = f'''
            SELECT 
            SUM( WYP.wyp_ZUSPrac ) AS ZUS_PRAC
            FROM plb_Wyplata WYP
            INNER JOIN plb_ListaPlac LP ON WYP.wyp_IdListyPlac = LP.lp_Id 
            INNER JOIN plb_Umowa UP ON UP.up_Id = WYP.wyp_IdUmowy
            INNER JOIN pr_Pracownik PR ON PR.pr_Id = UP.up_IdPracownika 
            WHERE (MONTH(LP.lp_Miesiac) = {self.month} AND YEAR(LP.lp_Miesiac) = {self.year}) AND lp_IdDefinicjiLP = 5
            GROUP BY LP.lp_Id, LP.lp_Miesiac, LP.lp_DataWyplaty
            ORDER BY LP.lp_Miesiac, LP.lp_DataWyplaty
        '''
        self.cursor.execute(query)
        zus_obcy_zlecenia = self.cursor.fetchone()
        if zus_obcy_zlecenia is None:
            zus_obcy_zlecenia = 0
        else:
            zus_obcy_zlecenia = zus_obcy_zlecenia[0]

        self.rows.append(
            (
                '229-051', as_curr(), as_curr(podatki1[2] - zus_obcy_zlecenia), 'Potrąc. ubezp. społ.',
                'Potrąc. ubezp. społ. - ZUS pracownik OBCY', ''
            )
        )

        self.rows.append(
            (
                '229-051', as_curr(), as_curr(zus_obcy_zlecenia), 'Potrąc. ubezp. społ. OBCY (ZLECENIA)',
                'ZUS pracownik OBCY', ''
            )
        )

        self.rows.append(
            (
                '229-051', as_curr(), as_curr(zus_pracownik_pracodawca_kwota_fp[0]),
                'Potrąc. ubezp. społ. Rada osoba', 'Potrąc. ubezp. społ.', ''
            )
        )

        self.rows.append(
            (
                '229-052', as_curr(), as_curr(podatki1[4] + podatki2[3]), 'Potrąc. ubezp. zdrow.',
                'Potrąc. ubezp. zdrow. 9%', ''
            )
        )

        self.rows.append(
            (
                '220-003', as_curr(), as_curr(podatki1[3] + podatki2[2]), 'Potrącenie zdrow.',
                'Potrącenie zdrow. 7,75%', ''
            )
        )

        self.rows.append(
            (
                '220-003', as_curr(podatki1[3] + podatki2[2]), as_curr(), 'Potrącenie zdrow.',
                'Potrącenie zdrow. 7,75%', ''
            )
        )

        self.rows.append(
            (
                '231-000', as_curr(potracenia_razem + zus_pracownik_pracodawca_kwota_fp[0]), as_curr(),
                'Razem potrącenia + dodatkowe potrącenia + Potrąc. ubezp. społ. Rada osoba',
                'Razem potrącenia podatek od bonów', ''
            )
        )

    def execute_queries(self):
        self.rows = []

        obcy_zlecenia = self.first_query()
        self.second_query(obcy_zlecenia)
        self.third_query(obcy_zlecenia)

        self.fourth_query()
        zus_pracownik_pracodawca_kwota_fp = self.fifth_query()
        self.sixth_query(zus_pracownik_pracodawca_kwota_fp)
        self.seventh_query(zus_pracownik_pracodawca_kwota_fp)
        self.eighth_query()
        self.ninth_query()

        funduszm_id = self.tenth_query()
        self.eleventh_query(funduszm_id)

        funduszm_razem = self.twelfth_query(funduszm_id)
        potracenia = self.thirteenth_query()
        pzu_id = self.fourteenth_query()
        pzu_razem = self.fifteenth_query(pzu_id)

        ubezp_szpitalne_id = self.sixteenth_query()
        ubezp_szpitalne_kwota = self.seventeenth_query(ubezp_szpitalne_id)

        splata_kzp_id = self.eighteenth_query()
        splata_kzp_kwota = self.nineteenth_query(splata_kzp_id)

        wklady_kzp_id = self.twentyth_query()
        wklady_kzp_kwota = self.twentyfirst_query(wklady_kzp_id)
        self.twentysecond_query(splata_kzp_kwota, wklady_kzp_kwota)

        podatki1 = self.twentythird_query()
        podatki2 = self.twentyfourth_query()
        potracenia_razem = funduszm_razem + potracenia + pzu_razem + ubezp_szpitalne_kwota + splata_kzp_kwota + \
            wklady_kzp_kwota + podatki1[0] + podatki2[0] + podatki1[1] + podatki2[1] + podatki1[2] + podatki1[4] + \
            podatki2[3]

        self.twentyfifth_query(podatki1, podatki2, potracenia_razem, zus_pracownik_pracodawca_kwota_fp)
