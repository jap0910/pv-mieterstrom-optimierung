#
# Wirtschaftlichkeitsberechnung: PV-Erzeugung vs. Verbrauch 
# ===============================================================

#
# 0) Ausgangsdaten & Variablen festlegen
# --------------------------------------------------------------
# Projektinformationen
    projekt = "Leipzig"       # beispielhafte (Orts-)Bezeichnung zur Dateibezeichnung
    sunYear = "2020"          # betrachtetes Sonnenjahr bei der PV-Stromproduktion
    kWp = 20                  # Nennleistung der PV-Anlage in kWp
    laufzeit = 20             # Laufzeit des Projekts

# Anzahl Smart-Meter in der Liegenschaft (Grenzen der Optimierung; Wärmepumpe (WP) an die Anzahl der Wohneinheiten (WE) gekoppelt)
    if kWp == 10
        alle_WE = 5
        ALLG_mx = 2
        GEW_mx = 1
    elseif kWp == 20
        alle_WE = 10
        ALLG_mx = 2
        GEW_mx = 1
    elseif kWp == 30
        alle_WE = 20
        ALLG_mx = 3
        GEW_mx = 2
    else
        error("Kein Anzahl an Wohneinheiten für $(kWp) kWp hinterlegt — bitte ergänzen.")
        exit()
    end
    limits = Dict(
        :MIE_min =>0, :MIE_max => alle_WE, :MIE_step => 1.0,
        :ALLG_min => 0, :ALLG_max => ALLG_mx, :ALLG_step => 1.0,
        :GEW_min => 0, :GEW_max => GEW_mx, :GEW_step => 1.0,
        :EMOB_min => 0, :EMOB_max => 1, :EMOB_step => 1.0
    )

# Speicherparameter
    speicher_kapazitaet = 11      # Nennkapazität (kWh)
    speicher_soc_init = 0         # initialer Ladezustand
    speicher_wirkungsgrad_laden = 0.95      # Wirkungsgrad der Ladeleistung
    speicher_wirkungsgrad_entladen = 0.95   # Wirkungsgrad der Entladeleistung
    speicher_leistung_max = 5 * 0.25        # max. Lade-/ Entladeleistung (kWh)

# Investitionskosten
    pv_invest = 1000 * kWp              # für PV-Anlage 
    speicher_invest = (550 * speicher_kapazitaet) + 3000   # für Batteriespeicher
    emob_invest = 2000                  # je Elektro-Ladesäule
    wp_invest = 5000 * alle_WE + 5000   # für Wärmepumpe
    zähler_invest = 100                 # je Zähler

# Umsatzvariablen
    # Grundversorgertarif
    preis_gvABP = 0.2593     # bspw. Leipzig: 0.2593 , München: 0.2726 , Hamburg: 0.3231 
    preis_gvGP = 14.78       # bspw. Leipzig: 14.78 , München: 11,31 , Hamburg: 15.04
    # Mieterstromtarif je Teilnehmergruppe (& für PV- und Netzstrom für opt. 2-Tarif-Modell)
    preis_mieABP_son = preis_gvABP * 0.85   # Mieterstromtarif 85 % des Grundversorgertarifs (Arbeitspreis)
    preis_mieABP_netz = preis_gvABP * 0.85  # (Grundpreis)
    preis_mieGP = preis_gvGP * 0.85
    preis_allgABP = preis_gvABP * 0.85
    preis_allgGP = preis_gvGP * 0.85
    preis_gewABP = preis_gvABP * 0.85
    preis_gewGP = preis_gvGP * 0.85
    preis_emobABP = preis_gvABP * 0.85
    preis_emobGP = preis_gvGP * 0.85
    preis_wpABP = preis_gvABP * 0.85
    if kWp == 10                # Vergütungssätze für eine PV-Anlage mit 10 kWp
        verg_mie    = 0.0256    # Mieterstromzuschlag
        verg_eeg    = 0.0786    # Teileinspeisevergütung
        verg_eeg_VE = 0.1247    # Volleinspeisevergütung
    elseif kWp == 20            # Vergütungssätze für eine PV-Anlage mit 20 kWp
        verg_mie    = 0.0247
        verg_eeg    = 0.0733
        verg_eeg_VE = 0.1146
    elseif kWp == 30            # Vergütungssätze für eine PV-Anlage mit 30 kWp
        verg_mie    = 0.0244
        verg_eeg    = 0.0715
        verg_eeg_VE = 0.1112
    else
        error("Kein Vergütungssatz für $(kWp) kWp hinterlegt — bitte ergänzen.")
        exit()
    end
    pog_uml = 16.81         # Umlagefähige Kosten für Betrieb Smart-Meter

# Kostenvariablen
    smart_met = 30          # Kosten je Smart-Meter pro Jahr
    cost_ems = 5            # Kosten für Energiemanagement
    pv_wart = 0.01 * pv_invest          # Wartung PV-Anlage
    speicher_wart = 0.01 * speicher_invest     # Wartung Speicher
    emob_wart = 0.01 * emob_invest      # Wartung Elektro-Ladesäule
    wp_wart = 0.01 * wp_invest          # Wartung Wärmepumpe
    preis_netzABP = preis_gvABP * 0.95  # Reststromtarif 85 % des Grundversorgertarifs (Arbeitspreis)
    preis_netzGP = preis_gvGP * 0.95    # (Grundpreis)

# Notwendige Pakete aufrufen
    using DataFrames
    using XLSX
    using Dates
    using Statistics
    using Plots
    using StatsPlots
    using CSV
    using Printf
    using FinanceCore    

println("Projektdaten wurden konfiguriert; starte Einlesen von Daten...")

# 
# 1) Datengrundlage einlesen
# ---------------------------------------------------------------
# Dateipfade & -namen von Inputdaten definieren
    # Ordnerpfad
    base_path = @__DIR__

    # Produktion einer PV-Anlage
    file_name_PV = joinpath(base_path, "..", "data", "pv", "$(projekt)_PV-Power_$(kWp)kWp_$(sunYear).xlsx")

    # Standardlastprofile aufrufen
    file_name_load_MIE  = joinpath(base_path, "..", "data", "load_profiles", "Energy-Demand_MIE-SLP.xlsx")
    file_name_load_ALLG  = joinpath(base_path,  "..", "data", "load_profiles", "Energy-Demand_ALLG-Aufzug.xlsx")
    file_name_load_GEW  = joinpath(base_path,  "..", "data", "load_profiles", "Energy-Demand_GEW-SLP-klein.xlsx")
    file_name_load_EMOB  = joinpath(base_path,  "..", "data", "load_profiles", "Energy-Demand_EMOB-SLP-2250kWh.xlsx")
    file_name_load_WP  = joinpath(base_path,  "..", "data", "load_profiles", "Energy-Demand_WP-SLP.xlsx")

    # Funktion zur Angleichung der geladenen Daten
    function lade_daten(dateipfad)
        df = XLSX.readtable(dateipfad, 1) |> DataFrame                  # erstes Tabellenblatt wird als DataFrame eingelesen
        rename!(df, Dict("Timestamp" => "Zeit"))                        # Vereinheitlichung Spaltenname für Zeit 
        df.Zeit = DateTime.(df.Zeit, dateformat"yyyy-mm-ddTHH:MM:SS")   # Wandelt Strings aus Excel in DateTime-Objekte um
        sort!(df, :Zeit)         # Daten nach Zeit sortieren
        return df
    end

# Daten laden
    df_PV        = lade_daten(file_name_PV)
    df_load_MIE  = lade_daten(file_name_load_MIE)
    df_load_ALLG = lade_daten(file_name_load_ALLG)
    df_load_GEW  = lade_daten(file_name_load_GEW)
    df_load_EMOB = lade_daten(file_name_load_EMOB)
    df_load_WP   = lade_daten(file_name_load_WP)

# PV- und Verbrauchsdaten zusammenführen mit Hilfsspalte "Zeit_ohne_Jahr"
    for df in [df_PV, df_load_MIE, df_load_ALLG, df_load_GEW, df_load_EMOB, df_load_WP]
        df[!, :Zeit_ohne_Jahr] = DateTime.(2020, month.(df.Zeit), day.(df.Zeit), hour.(df.Zeit), minute.(df.Zeit))
    end

    sheet_data = innerjoin(df_PV, df_load_MIE, on=:Zeit_ohne_Jahr, makeunique=true)
    for df in [df_load_ALLG, df_load_GEW, df_load_EMOB, df_load_WP]
        global sheet_data = innerjoin(sheet_data, df, on=:Zeit_ohne_Jahr, makeunique=true)
    end

# Eingelesene Daten formatieren
    # Relevante Spalten behalten 
    select!(sheet_data, [
        "Zeit",
        "PV-Power [kWh]",
        "Energy-Demand [kWh]",
        "Energy-Demand [kWh]_1",
        "Energy-Demand [kWh]_2",
        "Energy-Demand [kWh]_3",
        "Energy-Demand [kWh]_4"
    ])

    # Einheitliche, kurze Spaltennamen vergeben
    rename!(sheet_data, Dict(
        "PV-Power [kWh]" => "PV-Stromproduktion [kWh]",
        "Energy-Demand [kWh]" => "Stromverbrauch MIE [kWh]",
        "Energy-Demand [kWh]_1" => "Stromverbrauch ALLG [kWh]",
        "Energy-Demand [kWh]_2" => "Stromverbrauch GEW [kWh]",
        "Energy-Demand [kWh]_3" => "Stromverbrauch EMOB [kWh]",
        "Energy-Demand [kWh]_4" => "Stromverbrauch WP [kWh]"
    ))

    # Gesamtverbrauch in kWh berechnen
    sheet_data[!,"Gesamtstromverbrauch [kWh]"] =
        sheet_data[!,"Stromverbrauch MIE [kWh]"] .+
        sheet_data[!,"Stromverbrauch ALLG [kWh]"] .+
        sheet_data[!,"Stromverbrauch GEW [kWh]"] .+
        sheet_data[!,"Stromverbrauch EMOB [kWh]"] .+
        sheet_data[!,"Stromverbrauch WP [kWh]"]

    # Monate als Namen formatieren
    const MONTH_NAMES_DE = [
        "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember"
    ]
    monthname_de(m::Int) = MONTH_NAMES_DE[m]

println()
println("Inputdaten erfolgreich eingelesen, zusammengeführt & formatiert! Insgesamt $(nrow(sheet_data)) Zeilen.")

#
# 3. Bilanzierung
# --------------------------------------------------------------
# Verbraucherprofile auf gegebene Teilnehmer:innenzahlen skalieren
    function scale_loads(sheet_data::DataFrame,
                        nMIE::Real, nALLG::Real, nGEW::Real, nEMOB::Real, nWP::Real)

        df = deepcopy(sheet_data) # Kopie erstellen, damit keine Daten überschrieben werden

        df[!,"Stromverbrauch MIE [kWh]"]   .= sheet_data[!,"Stromverbrauch MIE [kWh]"]   .* nMIE
        df[!,"Stromverbrauch ALLG [kWh]"]  .= sheet_data[!,"Stromverbrauch ALLG [kWh]"]  .* nALLG
        df[!,"Stromverbrauch GEW [kWh]"]   .= sheet_data[!,"Stromverbrauch GEW [kWh]"]   .* nGEW
        df[!,"Stromverbrauch EMOB [kWh]"]  .= sheet_data[!,"Stromverbrauch EMOB [kWh]"]  .* nEMOB
        df[!,"Stromverbrauch WP [kWh]"]    .= sheet_data[!,"Stromverbrauch WP [kWh]"]    .* nWP

        df[!,"Gesamtstromverbrauch [kWh]"] .=
            df[!,"Stromverbrauch MIE [kWh]"] .+
            df[!,"Stromverbrauch ALLG [kWh]"] .+
            df[!,"Stromverbrauch GEW [kWh]"] .+
            df[!,"Stromverbrauch EMOB [kWh]"] .+
            df[!,"Stromverbrauch WP [kWh]"]

        return df
    end

# Strombilanz ohne Speicher (15-min)
    function simulate_no_storage(scaled::DataFrame)
        df = deepcopy(scaled)           # Datenkopie erstellen
        df[!,:Monat] = month.(df.Zeit)  # Monat für spätere Aggregation

        # Direkter Eigenverbrauch
        df[!,"PV-Stromverbrauch in-house [kWh]"] = map(eachrow(df)) do r
            min(r["PV-Stromproduktion [kWh]"], r["Gesamtstromverbrauch [kWh]"])
        end

        # Netzbezug
        df[!,"Netzbezug [kWh]"] = map(eachrow(df)) do r
            max(r["Gesamtstromverbrauch [kWh]"] - r["PV-Stromproduktion [kWh]"], 0.0)
        end

        # Einspeisung
        df[!,"PV-Einspeisung [kWh]"] = map(eachrow(df)) do r
            max(r["PV-Stromproduktion [kWh]"] - r["Gesamtstromverbrauch [kWh]"], 0.0)
        end

        # PV-Strom- & Netzbezugsanteile je Verbrauchergruppe
        gruppen = [
            "Stromverbrauch MIE [kWh]",
            "Stromverbrauch ALLG [kWh]",
            "Stromverbrauch GEW [kWh]",
            "Stromverbrauch EMOB [kWh]",
            "Stromverbrauch WP [kWh]"
        ]

        for g in gruppen
            pv_col   = replace(g, "[kWh]" => "PV_Anteil [kWh]")
            netz_col = replace(g, "[kWh]" => "Netzanteil [kWh]")

            df[!, pv_col] = map(eachrow(df)) do r
                G  = r["Gesamtstromverbrauch [kWh]"]
                EV = r["PV-Stromverbrauch in-house [kWh]"]
                V  = r[g]
                G <= 1e-9 ? 0.0 : (V / G) * EV
            end

            df[!, netz_col] = map(eachrow(df)) do r
                G  = r["Gesamtstromverbrauch [kWh]"]
                NB = r["Netzbezug [kWh]"]
                V  = r[g]
                G <= 1e-9 ? 0.0 : (V / G) * NB
            end
        end
        return df
    end

# Strombilanz mit Speicher inkl. Speicherfüllstand (15-min)
    function simulate_with_storage(scaled::DataFrame)
        df = deepcopy(scaled)               # Datenkopie erstellen
        df[!, :Monat] = month.(df.Zeit)     # Monat für spätere Aggregation

        soc = speicher_kapazitaet * speicher_soc_init

        n = nrow(df)
        speicherfuellstand = zeros(Float64, n)
        netzbezug = zeros(Float64, n)
        einspeisung = zeros(Float64, n)
        eigenverbrauch = zeros(Float64, n)

        for i in 1:n
            pv   = df[i, "PV-Stromproduktion [kWh]"]
            last = df[i, "Gesamtstromverbrauch [kWh]"]

            # Direkter Eigenverbrauch
            dirverbr = min(pv, last)
            rest_pv   = max(pv   - dirverbr, 0.0)
            rest_last = max(last - dirverbr, 0.0)

            # Laden
            lade = min(rest_pv, speicher_leistung_max, (speicher_kapazitaet - soc))
            soc += lade * speicher_wirkungsgrad_laden
            soc = min(soc, speicher_kapazitaet)

            # Entladen
            entlade_pot = min(rest_last, speicher_leistung_max)                 # max. nutzbare Energie durch Leistungslimit
            entnahme = min(entlade_pot / speicher_wirkungsgrad_entladen, soc)   # Speicher liefert nur, was er hat (unter Berücksichtigung Wirkungsgrad)
            genutzte_Energie = entnahme * speicher_wirkungsgrad_entladen        # Nutzbare Energie am Verbraucher
            soc -= entnahme     # Speicher aktualisieren
            soc = max(soc, 0.0)

            netzbezug[i]    = max(rest_last - genutzte_Energie, 0.0)
            einspeisung[i]  = max(rest_pv   - lade,    0.0)
            eigenverbrauch[i] = dirverbr + genutzte_Energie
            speicherfuellstand[i] = soc
        end

        df[!,"Netzbezug [kWh]"] = netzbezug
        df[!,"PV-Einspeisung [kWh]"] = einspeisung
        df[!,"PV-Stromverbrauch in-house [kWh]"] = eigenverbrauch
        df[!,"Speicherfüllstand [kWh]"] = speicherfuellstand

        # PV- & Netzbezugsanteile je Verbrauchergruppe
        gruppen = [
            "Stromverbrauch MIE [kWh]",
            "Stromverbrauch ALLG [kWh]",
            "Stromverbrauch GEW [kWh]",
            "Stromverbrauch EMOB [kWh]",
            "Stromverbrauch WP [kWh]"
        ]

        for g in gruppen
            pv_col   = replace(g, "[kWh]" => "PV_Anteil [kWh]")
            netz_col = replace(g, "[kWh]" => "Netzanteil [kWh]")

            df[!, pv_col] = map(eachrow(df)) do r
                G  = r["Gesamtstromverbrauch [kWh]"]
                EV = r["PV-Stromverbrauch in-house [kWh]"]
                V  = r[g]
                G <= 1e-9 ? 0.0 : (V / G) * EV
            end

            df[!, netz_col] = map(eachrow(df)) do r
                G  = r["Gesamtstromverbrauch [kWh]"]
                NB = r["Netzbezug [kWh]"]
                V  = r[g]
                G <= 1e-9 ? 0.0 : (V / G) * NB
            end
        end
        return df
    end

# Monatliche Aggregation der Daten 
    function monthly_aggregation(sim_df::DataFrame)
        monthly = combine(groupby(sim_df, :Monat),
            "Stromverbrauch MIE [kWh]"  => sum,
            "Stromverbrauch ALLG [kWh]" => sum,
            "Stromverbrauch GEW [kWh]"  => sum,
            "Stromverbrauch EMOB [kWh]" => sum,
            "Stromverbrauch WP [kWh]"   => sum,
            "Gesamtstromverbrauch [kWh]" => sum,
            "PV-Stromproduktion [kWh]"   => sum,
            "PV-Stromverbrauch in-house [kWh]"  => sum,
            "Netzbezug [kWh]"        => sum,
            "PV-Einspeisung [kWh]"   => sum,
            "Stromverbrauch MIE PV_Anteil [kWh]"   => sum,
            "Stromverbrauch MIE Netzanteil [kWh]"  => sum,
            "Stromverbrauch ALLG PV_Anteil [kWh]"  => sum,
            "Stromverbrauch ALLG Netzanteil [kWh]" => sum,
            "Stromverbrauch GEW PV_Anteil [kWh]"   => sum,
            "Stromverbrauch GEW Netzanteil [kWh]"  => sum,
            "Stromverbrauch EMOB PV_Anteil [kWh]"  => sum,
            "Stromverbrauch EMOB Netzanteil [kWh]" => sum,
            "Stromverbrauch WP PV_Anteil [kWh]"    => sum,
            "Stromverbrauch WP Netzanteil [kWh]"   => sum
        )

        # Auto-Umbenennung durch Aufsummierung entfernen
        rename!(monthly, Dict(
            "Stromverbrauch MIE [kWh]_sum"  => "Stromverbrauch MIE [kWh]",
            "Stromverbrauch ALLG [kWh]_sum" => "Stromverbrauch ALLG [kWh]",
            "Stromverbrauch GEW [kWh]_sum"  => "Stromverbrauch GEW [kWh]",
            "Stromverbrauch EMOB [kWh]_sum" => "Stromverbrauch EMOB [kWh]",
            "Stromverbrauch WP [kWh]_sum"   => "Stromverbrauch WP [kWh]",
            "Gesamtstromverbrauch [kWh]_sum" => "Gesamtstromverbrauch [kWh]",
            "PV-Stromproduktion [kWh]_sum"   => "PV-Stromproduktion [kWh]",
            "PV-Stromverbrauch in-house [kWh]_sum"  => "PV-Stromverbrauch in-house [kWh]",
            "Netzbezug [kWh]_sum"        => "Netzbezug [kWh]",
            "PV-Einspeisung [kWh]_sum"   => "PV-Einspeisung [kWh]",
            "Stromverbrauch MIE PV_Anteil [kWh]_sum"   => "Stromverbrauch MIE PV_Anteil [kWh]",
            "Stromverbrauch MIE Netzanteil [kWh]_sum"  => "Stromverbrauch MIE Netzanteil [kWh]",
            "Stromverbrauch ALLG PV_Anteil [kWh]_sum"  => "Stromverbrauch ALLG PV_Anteil [kWh]",
            "Stromverbrauch ALLG Netzanteil [kWh]_sum" => "Stromverbrauch ALLG Netzanteil [kWh]",
            "Stromverbrauch GEW PV_Anteil [kWh]_sum"   => "Stromverbrauch GEW PV_Anteil [kWh]",
            "Stromverbrauch GEW Netzanteil [kWh]_sum"  => "Stromverbrauch GEW Netzanteil [kWh]",
            "Stromverbrauch EMOB PV_Anteil [kWh]_sum"  => "Stromverbrauch EMOB PV_Anteil [kWh]",
            "Stromverbrauch EMOB Netzanteil [kWh]_sum" => "Stromverbrauch EMOB Netzanteil [kWh]",
            "Stromverbrauch WP PV_Anteil [kWh]_sum"    => "Stromverbrauch WP PV_Anteil [kWh]",
            "Stromverbrauch WP Netzanteil [kWh]_sum"   => "Stromverbrauch WP Netzanteil [kWh]"
        ))

        # Monatliche Eigenvebrauchs- & Autarkiequoten berechnen
        monthly[!,"Eigenverbrauchsquote [%]"] =
            monthly[!,"PV-Stromverbrauch in-house [kWh]"] ./ monthly[!,"PV-Stromproduktion [kWh]"] .* 100
        monthly[!,"Autarkiequote [%]"] =
            monthly[!,"PV-Stromverbrauch in-house [kWh]"] ./ monthly[!,"Gesamtstromverbrauch [kWh]"] .* 100
        return monthly
    end

# Wirtschaftliche Parameter berechnen
    function economic_eval(monthly::DataFrame, nMIE::Real, nALLG::Real, nGEW::Real, nEMOB::Real, nWP::Real; with_storage::Bool)
        df = deepcopy(monthly)

        # Umsätze
        df[!,"Stromverkauf Mieterstrom [€]"] =
            (df[!,"Stromverbrauch MIE PV_Anteil [kWh]"] .* preis_mieABP_son) .+
            (df[!,"Stromverbrauch MIE Netzanteil [kWh]"] .* preis_mieABP_netz) .+
            (preis_mieGP .* nMIE)

        df[!,"Stromverkauf Hausverwaltung (ALLG) [€]"] =
            (df[!,"Stromverbrauch ALLG [kWh]"] .* preis_allgABP) .+ (preis_allgGP .* nALLG)

        df[!,"Stromverkauf Gewerbe [€]"] =
            (df[!,"Stromverbrauch GEW [kWh]"] .* preis_gewABP) .+ (preis_gewGP .* nGEW)

        df[!,"Stromverkauf E-Mobilität [€]"] =
            (df[!,"Stromverbrauch EMOB [kWh]"] .* preis_emobABP) .+ (preis_emobGP .* nEMOB)
        
        df[!,"Stromverkauf Wärme [€]"] =
            (df[!,"Stromverbrauch WP [kWh]"] .* preis_wpABP)

        df[!,"Vergütung Mieterstromzuschlag [€]"] =
            df[!,"PV-Stromverbrauch in-house [kWh]"] .* verg_mie

        df[!,"EEG-Vergütung [€]"] =
            df[!,"PV-Einspeisung [kWh]"] .* verg_eeg

        df[!,"Umlagefähige Smart-Meter-Kosten [€]"] .= (pog_uml ./ 12) .* (nMIE .+ nGEW)

        df[!,"Umsatz gesamt [€]"] =
            df[!,"Stromverkauf Mieterstrom [€]"] .+
            df[!,"Stromverkauf Hausverwaltung (ALLG) [€]"] .+
            df[!,"Stromverkauf Gewerbe [€]"] .+
            df[!,"Stromverkauf E-Mobilität [€]"] .+
            df[!,"Stromverkauf Wärme [€]"] .+
            df[!,"Vergütung Mieterstromzuschlag [€]"] .+
            df[!,"EEG-Vergütung [€]"] .+
            df[!,"Umlagefähige Smart-Meter-Kosten [€]"]

        # Kosten
        df[!,"Kosten Smart-Meter [€]"] .= (smart_met ./ 12) .* (nMIE .+ nALLG .+ nGEW .+ nEMOB)
        df[!,"Kosten EMS [€]"] .= cost_ems .* (nMIE .+ nALLG .+ nGEW .+ nEMOB .+ 4 + (nWP > 0 ? 1.0 : 0.0))
        df[!,"Wartung PV-System [€]"] .= pv_wart ./ 12

        if with_storage
            df[!,"Wartung Speicher [€]"] .= speicher_wart ./ 12
        else
            df[!,"Wartung Speicher [€]"] .= 0.0
        end

        df[!,"Wartung E-Mobilität [€]"] .= (emob_wart ./ 12) .* nEMOB
        df[!,"Wartung Wärmepumpe [€]"]  .= (wp_wart   ./ 12) .* (nWP > 0 ? 1.0 : 0.0)

        df[!,"Stromeinkauf Netzbezug [€]"] =
            df[!,"Netzbezug [kWh]"] .* preis_netzABP .+ preis_netzGP

        df[!,"Kosten gesamt [€]"] =
            df[!,"Kosten Smart-Meter [€]"] .+
            df[!,"Kosten EMS [€]"] .+
            df[!,"Wartung PV-System [€]"] .+
            df[!,"Wartung Speicher [€]"] .+
            df[!,"Wartung E-Mobilität [€]"] .+
            df[!,"Wartung Wärmepumpe [€]"] .+
            df[!,"Stromeinkauf Netzbezug [€]"]

        # Gewinn
        gewinn = sum(df[!,"Umsatz gesamt [€]"]) - sum(df[!,"Kosten gesamt [€]"])
        return gewinn, df
    end

# Berechnung & Abwägunng, ob ohne vs. mit Speicher
    function evaluate_case(sheet_data::DataFrame, nMIE::Real, nALLG::Real, nGEW::Real, nEMOB::Real, nWP::Real; with_storage::Bool)

        scaled = scale_loads(sheet_data, nMIE, nALLG, nGEW, nEMOB, nWP) # Verbrauchsprofile auf gewünschte Anzahl Verbraucher skalieren
        sim_df = with_storage ? simulate_with_storage(scaled) : simulate_no_storage(scaled) # Simulation je nach Speicher-Option
        monthly = monthly_aggregation(sim_df) # Aggregation in Monatswerte

        gewinn, _ = economic_eval(monthly, nMIE, nALLG, nGEW, nEMOB, nWP; with_storage=with_storage) # Netto-Gewinn für Szenario
        return gewinn
    end

# Referenz: PV-Volleinspeisung
    scaled_VE  = scale_loads(sheet_data, 0, 0, 0, 0, 0)
    sim_VE     = simulate_no_storage(scaled_VE)
    monthly_VE = monthly_aggregation(sim_VE)

    gewinn_VE_jahr = sum(monthly_VE[!, "PV-Einspeisung [kWh]"]) * verg_eeg_VE - pv_wart

println()
println("Bilanzierung erfolgreich durchgelaufen!")

#
# 4) Optimierung
# --------------------------------------------------------------
# Limits der Verbrauchsgruppen
    rangeMIE   = limits[:MIE_min]:limits[:MIE_step]:limits[:MIE_max]
    rangeALLG  = limits[:ALLG_min]:limits[:ALLG_step]:limits[:ALLG_max]
    rangeGEW   = limits[:GEW_min]:limits[:GEW_step]:limits[:GEW_max]
    rangeEMOB  = limits[:EMOB_min]:limits[:EMOB_step]:limits[:EMOB_max]

# Optimales Ergebnis ohne Speicher
    ergebnisse_oSP = DataFrame(nMIE=Float64[], nALLG=Float64[], nGEW=Float64[],
                            nEMOB=Float64[], nWP=Float64[], Gewinn_oSP=Float64[])

    for nMIE in rangeMIE, nALLG in rangeALLG, nGEW in rangeGEW, nEMOB in rangeEMOB
        if nMIE + nGEW <= alle_WE
            nWP_A = alle_WE     # Option A: alle WE der Liegenschaft 
            gA = evaluate_case(sheet_data, nMIE, nALLG, nGEW, nEMOB, nWP_A; with_storage=false)
            push!(ergebnisse_oSP, (nMIE, nALLG, nGEW, nEMOB, nWP_A, gA))

            nWP_B = 0.0         # Option B: keine WP
            gB = evaluate_case(sheet_data, nMIE, nALLG, nGEW, nEMOB, nWP_B; with_storage=false)
            push!(ergebnisse_oSP, (nMIE, nALLG, nGEW, nEMOB, nWP_B, gB))
        end
    end

    best_idx_oSP   = argmax(ergebnisse_oSP.Gewinn_oSP)
    best_combo_oSP = ergebnisse_oSP[best_idx_oSP, :]
    
    # Investitionskosten und Amortisationszeit
    wp_invest_oSP = best_combo_oSP.nWP > 0 ? wp_invest : 0.0
    invest_oSP = ((zähler_invest * (best_combo_oSP.nMIE + best_combo_oSP.nALLG + best_combo_oSP.nGEW + best_combo_oSP.nEMOB))
            + pv_invest 
            + (emob_invest * best_combo_oSP.nEMOB) 
            + (wp_invest_oSP))
    amozeit_oSP = invest_oSP / best_combo_oSP.Gewinn_oSP

    # IRR (Internal-Reate-of-Investment)
    cashflows_oSP = vcat(-invest_oSP, fill(best_combo_oSP.Gewinn_oSP, laufzeit))
    irr_oSP = best_combo_oSP.Gewinn_oSP > 0 ? rate(irr(cashflows_oSP)) : NaN

# Optimales Ergebnis mit Speicher
    ergebnisse_mSP = DataFrame(nMIE=Float64[], nALLG=Float64[], nGEW=Float64[],
                            nEMOB=Float64[], nWP=Float64[], Gewinn_mSP=Float64[])

    for nMIE in rangeMIE, nALLG in rangeALLG, nGEW in rangeGEW, nEMOB in rangeEMOB
        if nMIE + nGEW <= alle_WE
            nWP_A = alle_WE     # Option A: alle WE der Liegenschaft 
            gA = evaluate_case(sheet_data, nMIE, nALLG, nGEW, nEMOB, nWP_A; with_storage=true)
            push!(ergebnisse_mSP, (nMIE, nALLG, nGEW, nEMOB, nWP_A, gA))

            nWP_B = 0.0         # Option B: keine WP
            gB = evaluate_case(sheet_data, nMIE, nALLG, nGEW, nEMOB, nWP_B; with_storage=true)
            push!(ergebnisse_mSP, (nMIE, nALLG, nGEW, nEMOB, nWP_B, gB))
        end
    end

    best_idx_mSP   = argmax(ergebnisse_mSP.Gewinn_mSP)
    best_combo_mSP = ergebnisse_mSP[best_idx_mSP, :]

    # Investitionskosten und Amortisationszeit
    wp_invest_mSP = best_combo_mSP.nWP > 0 ? wp_invest : 0.0
    invest_mSP = ((zähler_invest * (best_combo_mSP.nMIE + best_combo_mSP.nALLG + best_combo_mSP.nGEW + best_combo_mSP.nEMOB))
                + pv_invest 
                + speicher_invest 
                + (emob_invest * best_combo_mSP.nEMOB) 
                + (wp_invest_mSP))
    amozeit_mSP = invest_mSP / best_combo_mSP.Gewinn_mSP

    # IRR (Internal-Reate-of-Investment)
    cashflows_mSP = vcat(-invest_mSP, fill(best_combo_mSP.Gewinn_mSP, laufzeit))
    irr_mSP = best_combo_mSP.Gewinn_mSP > 0 ? rate(irr(cashflows_mSP)) : NaN

    # PV-LCOE (Stromgestehungskosten)
    lcoe = (pv_invest + laufzeit * pv_wart) / (laufzeit * sum(sheet_data[!, "PV-Stromproduktion [kWh]"]))

# Auswahl nach maximalen Jahresgewinn
    best_gewinn_mSP = best_combo_mSP.Gewinn_mSP
    best_gewinn_oSP = best_combo_oSP.Gewinn_oSP

    println()
    println("Die Optimierung des Mieterstromprojekts ist abgeschlossen. Die Berechnungen kommen zu folgendem Ergebnis:")
    println(@sprintf("Referenz Volleinspeisung (EEG) = %.2f €/Jahr", gewinn_VE_jahr))
    println(@sprintf("LCOE_PV = %.2f Cent/kWh", lcoe*100))
    println()

    if gewinn_VE_jahr >= best_gewinn_mSP && gewinn_VE_jahr >= best_gewinn_oSP
        # Optimum: Volleinspeisung 
        nMIE_opt  = 0.0; nALLG_opt = 0.0; nGEW_opt  = 0.0
        nEMOB_opt = 0.0; nWP_opt   = 0.0
        use_storage = false

        println()
        println("Die Volleinspeisung ins Netz erzielt den höchsten Jahresgewinn; kein Mieterstromkonzept empfohlen.")
        println(@sprintf("Differenz zu bester Konfig. mit SP  = %.2f €", gewinn_VE_jahr - best_gewinn_mSP))
        println(@sprintf("Differenz zu bester Konfig. ohne SP = %.2f €", gewinn_VE_jahr - best_gewinn_oSP))

    elseif best_gewinn_mSP >= best_gewinn_oSP
        # Optimum: Mieterstrom mit Speicher
        use_storage = true
        nMIE_opt  = best_combo_mSP.nMIE
        nALLG_opt = best_combo_mSP.nALLG
        nGEW_opt  = best_combo_mSP.nGEW
        nEMOB_opt = best_combo_mSP.nEMOB
        nWP_opt   = best_combo_mSP.nWP

        println()
        println("Der höchste Jahresgewinn wird erzielt bei:
                Mieter bzw. private Haushalte (nMIE)  = $(nMIE_opt)
                Allgemeinstromzähler (nALLG)           = $(nALLG_opt)
                Gewerbeeinheiten (nGEW)                = $(nGEW_opt)
                Elektroladesäulen (nEMOB)              = $(nEMOB_opt)
                Wärmepumpe (nWP)                       = $(nWP_opt)
                Batteriespeicher (mSP)                 = ja")
        println()
        println(@sprintf("Maximaler jährlicher Gewinn         = %.2f €", best_gewinn_mSP))
        println(@sprintf("Differenz Gewinn vs. Volleinspeisung = %.2f €", best_gewinn_mSP - gewinn_VE_jahr))
        println(@sprintf("Amortisationszeit                   = %.2f Jahre", amozeit_mSP))
        println(@sprintf("IRR                                 = %.2f Prozent", irr_mSP*100))
        println()
        println("Im Vergleich zu ohne Speicher gilt:")
        println(@sprintf("Differenz Gewinn (mSP - oSP)        = %.2f €", best_gewinn_mSP - best_gewinn_oSP))
        println(@sprintf("Differenz Amortisationszeit         = %.2f Jahre", amozeit_mSP - amozeit_oSP))
        println(@sprintf("IRR ohne Speicher                   = %.2f Prozent", irr_oSP*100))

    else
        # Optimum: Mieterstrom ohne Speicher
        use_storage = false
        nMIE_opt  = best_combo_oSP.nMIE
        nALLG_opt = best_combo_oSP.nALLG
        nGEW_opt  = best_combo_oSP.nGEW
        nEMOB_opt = best_combo_oSP.nEMOB
        nWP_opt   = best_combo_oSP.nWP

        println("Der höchste Jahresgewinn wird erzielt bei:
                Mieter bzw. private Haushalte (nMIE)  = $(nMIE_opt)
                Allgemeinstromzähler (nALLG)           = $(nALLG_opt)
                Gewerbeeinheiten (nGEW)                = $(nGEW_opt)
                Elektroladesäulen (nEMOB)              = $(nEMOB_opt)
                Wärmepumpe (nWP)                       = $(nWP_opt)
                Batteriespeicher (mSP)                 = nein")
        println()
        println(@sprintf("Maximaler jährlicher Gewinn         = %.2f €", best_gewinn_oSP))
        println(@sprintf("Differenz Gewinn vs. Volleinspeisung = %.2f €", best_gewinn_oSP - gewinn_VE_jahr))
        println(@sprintf("Amortisationszeit                   = %.2f Jahre", amozeit_oSP))
        println(@sprintf("IRR                                 = %.2f Prozent", irr_oSP*100))
    end
    
# Beenden des Codes, wenn Optimum Volleinspeisung ist
    if nMIE_opt + nALLG_opt + nGEW_opt + nEMOB_opt + nWP_opt == 0
        println()
        println("Volleinspeisung ist optimal; es werden keine Excel-Datei und keine Plots erstellt.")
        exit()
    end

println()
println("Die Ergebnisse können jetzt in einer Excel-Datei gespeichert werden.")

#
# 5) Datenablage in Excel-Datei
# ---------------------------------------------------------------
# Simulation ohne Speicher für Excel-Export
    scaled_no = scale_loads(sheet_data, nMIE_opt, nALLG_opt, nGEW_opt, nEMOB_opt, nWP_opt)
    sim_no = simulate_no_storage(scaled_no)
    monthly_no = monthly_aggregation(sim_no)

    # Jährlicher Gewinn ohne Speicher
    gewinn_no, monthly_no = economic_eval(monthly_no, nMIE_opt, nALLG_opt, nGEW_opt, nEMOB_opt, nWP_opt; 
                            with_storage=false)

# Simulation mit Speicher für Excel-Export
    scaled_yes = scale_loads(sheet_data, nMIE_opt, nALLG_opt, nGEW_opt, nEMOB_opt, nWP_opt)
    sim_yes = simulate_with_storage(scaled_yes)
    monthly_yes = monthly_aggregation(sim_yes)

    # Jährlicher Gewinn mit Speicher
    gewinn_yes, monthly_yes = economic_eval(monthly_yes, nMIE_opt, nALLG_opt, nGEW_opt, nEMOB_opt, nWP_opt; 
                            with_storage=true)

# Jahreswerte berechnen
    function build_summary(monthly::DataFrame)
        jahr = DataFrame(Monat = ["Gesamtes Jahr"]) # neue Spalte für Jahreswerte

        for col in names(monthly) # Iteration über alle Spalten mit Monatsnamen
            if col == "Monat" 
                continue
            end

            vals = monthly[!, col]
            scol = Symbol(col) 

            if eltype(vals) <: Number
                if col in ["Eigenverbrauchsquote [%]", "Autarkiequote [%]"]
                    if col == "Eigenverbrauchsquote [%]"
                        jahr[!, scol] = [sum(monthly[!,"PV-Stromverbrauch in-house [kWh]"]) /
                            sum(monthly[!,"PV-Stromproduktion [kWh]"]) * 100]
                    elseif col == "Autarkiequote [%]"
                        jahr[!, scol] = [sum(monthly[!,"PV-Stromverbrauch in-house [kWh]"]) /
                            sum(monthly[!,"Gesamtstromverbrauch [kWh]"]) * 100]
                    end
                else
                    jahr[!, scol] = [sum(skipmissing(vals))] 
                end
            else
                jahr[!, scol] = [missing]
            end
        end

        summary = vcat(monthly, jahr; cols = :union)
        # je Spalte Gewinn hinzufügen
        summary[!, "Gewinn [€]"] =
            summary[!, "Umsatz gesamt [€]"] .-
            summary[!, "Kosten gesamt [€]"]

        return summary
    end

    # Auswertung für Optimum ohne vs. mit Speicher
    summary_no = build_summary(monthly_no)
    summary_yes = build_summary(monthly_yes)

# Tabelle transponieren
    function make_transposed(df::DataFrame) 
        data_part = select(df, Not(:Monat)) 
        data_matrix = permutedims(Matrix(data_part)) 
        
        month_labels = vcat(MONTH_NAMES_DE, ["Gesamtes Jahr"]) # Monatsnamen erzeugen 
        t = DataFrame(data_matrix, Symbol.(month_labels)) 
        insertcols!(t, 1, :Zeitraum => names(data_part)) 

        # Leerzeilen einfügen für Übersichtlichkeit
        empty_row = Any[""; fill(0, ncol(t)-1)...]
            insert!(t, 11, empty_row)
            insert!(t, 24, empty_row)
            insert!(t, 34, empty_row)
            insert!(t, 43, empty_row)

        return t 
    end

# Tabelle formatieren & schreiben
    function write_table(sheet, df::DataFrame)
        # Header
        for (j, name) in enumerate(names(df))
            sheet[1, j] = name
        end

        # Inhalte
        for i in 1:nrow(df), j in 1:ncol(df)
            val = df[i, j]
            sheet[i + 1, j] =
                (val isa Number && !ismissing(val)) ? round(val, digits=2) : val
        end
    end

# Dateipfad für Ablage des Exports
    # Zielordner definieren
    output_path = joinpath(base_path,  "..", "results", "$(projekt)_Ergebnisse")

    # neuen Ordner erstellen (nur wenn nicht vorhanden)
    if !isdir(output_path)
        mkpath(output_path)
    end

    # Dateipfad
    output_file = joinpath(output_path, "$(projekt)_Ergebnisübersicht.xlsx") 

# Excel-Export
    XLSX.openxlsx(output_file, mode="w") do xf
        # Erstes Sheet anlegen als „Ohne Speicher“
        sheet1 = xf[1]
        sheet1.name = "Ohne Speicher"
        write_table(sheet1, make_transposed(summary_no))

        sheet1[47, 1] = "Investitionskosten [€]"
        sheet1[47, 2] = invest_oSP          
        sheet1[48, 1] = "Amortisationszeit [Jahre]"
        sheet1[48, 2] = amozeit_oSP        
        sheet1[49, 1] = "IRR [Prozent]"
        sheet1[49, 2] = irr_oSP 
        
        sheet1[51, 1] = "Vgl. Volleinspeisung mit EEG-Vergütung"
        sheet1[51, 2] = (sum(sim_no[!, "PV-Stromproduktion [kWh]"]) * verg_eeg_VE) - pv_wart
        sheet1[52, 1] = "PV-LCOE [Euro]"
        sheet1[52, 2] = lcoe
        
        # Zweites Sheet hinzufügen als „Mit Speicher“
        sheet2 = XLSX.addsheet!(xf, "Mit Speicher")
        write_table(sheet2, make_transposed(summary_yes))

        sheet2[47, 1] = "Investitionskosten [€]"
        sheet2[47, 2] = invest_mSP
        sheet2[48, 1] = "Amortisationszeit [Jahre]"
        sheet2[48, 2] = amozeit_mSP
        sheet2[49, 1] = "IRR [Prozent]"
        sheet2[49, 2] = irr_mSP        
        
        sheet2[51, 1] = "Vgl. Volleinspeisung mit EEG-Vergütung"
        sheet2[51, 2] = (sum(sim_no[!, "PV-Stromproduktion [kWh]"]) * verg_eeg_VE) - (pv_wart + speicher_wart)
        sheet2[52, 1] = "PV-LCOE [Euro]"
        sheet2[52, 2] = lcoe
    end

println()
println("Die Daten wurden erfolgreich in einer Excel-Datei gespeichert!")

#
# 6) Visualisierungen bzw. Plots erstellen
# ---------------------------------------------------------------
# Datengrundlage für Visualisierungen: mit Speicher als sim_yes, ohne Speicher als sim_no
    sim_opt = use_storage ? sim_yes : sim_no

# Monatsübersicht (15-min-Werte)
    months = 1:12
    for m in months
        df_month = filter(row -> month(row.Zeit) == m, sim_opt)
        if nrow(df_month) == 0
            continue
        end

        plt = bar(
            df_month.Zeit, df_month[!,"Gesamtstromverbrauch [kWh]"],
            label = "Gesamtstromverbrauch [kWh]",
            color = :red, linewidth = 0,
            fillrange = 0, fillalpha = 0.6,
            legend = :topright,
            xlabel = "Datum",
            ylabel = "Strom [kWh]",
            title = "PV-Stromproduktion vs. Stromverbrauch im $(monthname_de(m)) [15min-Werte]",
            left_margin=6Plots.mm,
            right_margin=3Plots.mm,
            bottom_margin=5Plots.mm,
            top_margin=3Plots.mm,
            size = (1000, 500),
            dpi = 500
        )
        bar!(plt, df_month.Zeit, df_month[!,"PV-Stromproduktion [kWh]"],
            label = "PV-Stromproduktion [kWh]",
            color = :gold, linewidth = 0,
            fillrange = 0, fillalpha = 0.8
        )

        # Speicherfüllstand nur plotten, wenn die Spalte existiert
        if "Speicherfüllstand [kWh]" in names(df_month)
            plot!(plt, df_month.Zeit, df_month[!,"Speicherfüllstand [kWh]"],
                label = "Speicherfüllstand [kWh]",
                color = :grey, linewidth = 1, linestyle = :dash
            )
        end

        xticks = collect(first(df_month.Zeit):Day(2):last(df_month.Zeit))
        xticklabels = Dates.format.(xticks, dateformat"dd.mm.")
        xticks!(plt, Dates.value.(xticks), xticklabels)

        savefig(plt, joinpath(output_path, "$(projekt)_Monat-$(lpad(m, 2, '0')).png"))
    end

# Beispielhafte Tage (15-min-Werte)
    example_days = [Date(2020,3,17)]
    for d in example_days
        start = d
        stop  = d + Day(2)   # Exemplarischer Tag plus 2 Tage
        df_day = filter(row -> start <= Date(row.Zeit) <= stop, sim_opt)
        if nrow(df_day) == 0
            continue
        end

        x = Dates.value.(df_day.Zeit)
        plt = bar(
            x, df_day[!,"Gesamtstromverbrauch [kWh]"],
            label = "Gesamtstromverbrauch [kWh]",
            color = :red,
            linewidth = 0, fillrange = 0, fillalpha = 0.6,
            legend = :topright,
            xlabel = "Uhrzeit",
            ylabel = "Strom [kWh]",
            title = "PV-Stromproduktion vs. Stromverbrauch vom $(Dates.format(d, "dd.")) bis $(Dates.format(d+Day(2), "dd.")) $(monthname_de(month(d))) [15min-Werte]",
            left_margin=6Plots.mm, right_margin=3Plots.mm,
            bottom_margin=5Plots.mm, top_margin=3Plots.mm,
            size = (1000, 500),
            dpi = 300
        )
        bar!(plt, x, df_day[!,"PV-Stromproduktion [kWh]"],
            label = "PV-Stromproduktion [kWh]",
            color = :gold, linewidth = 0,
            fillrange = 0, fillalpha = 0.7
        )

        if "Speicherfüllstand [kWh]" in names(df_day)
            plot!(plt, x, df_day[!,"Speicherfüllstand [kWh]"],
                label = "Speicherfüllstand [kWh]",
                color = :grey, linewidth = 1.5, linestyle = :dash
            )
        end

        xt = collect(first(df_day.Zeit):Hour(12):last(df_day.Zeit))
        xt_pos = Dates.value.(xt)
        xt_lab = [Dates.format(t, dateformat"dd.mm HH:MM") for t in xt]
        xticks!(plt, xt_pos, xt_lab)

        savefig(plt, joinpath(output_path, "$(projekt)_Tag-$(Dates.format(d, dateformat"mmdd"))_15min.png"))
    end

# Beispielhafte Tage (Stundenwerte, versch. Verbrauchsgruppen visualisiert)
    for d in example_days
        df_day = filter(row -> Date(row.Zeit) == d, sim_opt)
        df_day[!, :Stunde] = hour.(df_day.Zeit)

        # Aggregationen als Liste vordefinieren
        aggs = Any[
            "Stromverbrauch EMOB [kWh]" => sum => :EMOB_h,
            "Stromverbrauch MIE [kWh]"  => sum => :MIE_h,
            "Stromverbrauch ALLG [kWh]" => sum => :ALLG_h,
            "Stromverbrauch GEW [kWh]"  => sum => :GEW_h,
            "Stromverbrauch WP [kWh]"   => sum => :WP_h,
            "PV-Stromproduktion [kWh]"  => sum => :PV_h,
        ]
        
        if "Speicherfüllstand [kWh]" in names(df_day)
            push!(aggs, "Speicherfüllstand [kWh]" => last => :SP_h)
        end

        daily_stats = combine(
            groupby(df_day, :Stunde), aggs...
        )

        plt = groupedbar(
            daily_stats.Stunde,
            [daily_stats.EMOB_h daily_stats.WP_h daily_stats.MIE_h daily_stats.ALLG_h daily_stats.GEW_h],
            label = ["E-Ladesäulen" "Wärmepumpe" "Mieter" "Allgemeinstrom" "Gewerbe"],
            bar_position = :stack,
            color = [:cornflowerblue :mediumpurple3 :indianred :darkred :navyblue],
            linewidth = 0, fillrange = 0, fillalpha = 0.6,
            xlabel = "Stunde",
            ylabel = "Strom [kWh]",
            title = "PV-Stromproduktion vs. Stromverbrauch am $(day(d)). $(monthname_de(month(d))) [1h-Werte]",
            left_margin=6Plots.mm, right_margin=3Plots.mm,
            bottom_margin=5Plots.mm, top_margin=3Plots.mm,
            size = (1000, 500),
            dpi = 300
        )
        bar!(plt, daily_stats.Stunde, daily_stats.PV_h,
            label = "PV-Stromproduktion [kWh]",
            color = :gold, linewidth = 0, fillrange = 0, fillalpha = 0.7
        )
        if "Speicherfüllstand [kWh]" in names(df_day)
            plot!(plt, daily_stats.Stunde, daily_stats.SP_h,
                label="Speicherfüllstand [kWh]", 
                color=:grey, linewidth=1.5, linestyle=:dash
            )
        end

        savefig(plt, joinpath(output_path, "$(projekt)_Tag-$(Dates.format(d, dateformat"mmdd"))_1h.png"))
    end

# Jahresübersicht (Tageswerte)
    scaled = unique(sim_opt, :Zeit)
    scaled[!,:Tag] = Date.(scaled.Zeit)

    daily_stats2 = combine(
        groupby(scaled, :Tag),
        "Stromverbrauch EMOB [kWh]" => sum => :EMOB_Tag,
        "Stromverbrauch MIE [kWh]"  => sum => :MIE_Tag,
        "Stromverbrauch ALLG [kWh]" => sum => :ALLG_Tag,
        "Stromverbrauch GEW [kWh]"  => sum => :GEW_Tag,
        "Stromverbrauch WP [kWh]"   => sum => :WP_Tag,
        "PV-Stromproduktion [kWh]"  => sum => :PV_Tag
    )

    plt = groupedbar(daily_stats2.Tag, 
        [daily_stats2.EMOB_Tag daily_stats2.WP_Tag daily_stats2.MIE_Tag daily_stats2.ALLG_Tag daily_stats2.GEW_Tag],
        label = ["E-Ladesäulen [kWh]" "Wärmepumpe [kWh]" "Mieter [kWh]" "Allgemeinstrom [kWh]" "Gewerbe [kWh]"],
        bar_position = :stack,
        color = [:cornflowerblue :mediumpurple3 :indianred :darkred :navyblue], 
        linewidth = 0, fillrange = 0, fillalpha = 0.6,
        legend = :topright,
        xlabel = "Datum",
        ylabel = "Strom [kWh]",
        title = "Jahresübersicht: PV-Stromproduktion vs. Stromverbrauch [Tageswerte]",
        left_margin=6Plots.mm, right_margin=3Plots.mm,
        bottom_margin=5Plots.mm, top_margin=3Plots.mm,
        xrotation = 45,
        size = (1000, 500),
        dpi = 300
    )
    bar!(
        daily_stats2.Tag,
        daily_stats2.PV_Tag,
        label = "PV-Stromproduktion [kWh]",
        color = :gold, linewidth = 0, 
        fillrange = 0, fillalpha = 0.7,
    )

    xticks = collect(first(daily_stats2.Tag):Month(1):last(daily_stats2.Tag))
    xticklabels = [Dates.format(t, dateformat"dd.mm") for t in xticks]
    xticks!(plt, Dates.value.(xticks), xticklabels)

    savefig(plt, joinpath(output_path, "$(projekt)_Jahresübersicht-PVvsVerbrauch.png"))

println()
println("Plots wurden gespeichert.")