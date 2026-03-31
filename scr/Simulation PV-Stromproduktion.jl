#
# Simulation PV-Stromproduktion über ein Jahr
# ======================================

#
# 0) Ausgangsdaten & Variablen festlegen
# --------------------------------------------------------------
# Notwendige Pakete aufrufen
    using PythonCall
    using Dates
    using DataFrames

    pvlib = pyimport("pvlib")
    pd    = pyimport("pandas")
    np    = pyimport("numpy")
    pybuiltins = pyimport("builtins")

# Standortparamter und betrachtetes Jahr
    projekt = "Leipzig"         # Projektname bzw. Ortsbezeichnung
    lat, lon = 51.340, 12.374   # Standortkoordinaten, bspw. München: 48.138, 11.575 ; Leipzig: 51.340, 12.374 ; Hamburg: 53.550, 9.993 ; AA: 50.776, 6.083
    tz = "Europe/Berlin"        # Zeitzone
    sunYear = 2019              # Jahr der solaren Einstrahlung

# PV-Anlagen-Parameter
    tilt = 30.0                 # Neigung
    azim = 180.0                # Ausrichtung (180 = Süden)
    P_dc_W = 20_000.0           # Nennleistung der PV-Anlage in Watt-peak
    system_loss = 0.80          # pauschaler Systemwirkungsgrad der PV-Anlage
    n_inv = 0.98                # Wirkungsgrad Wechselrichter
    gamma_pdc = -0.004          # Temperaturkoeffizient der PV-Modulleistung

#
# 1) Verbindung zu PVGIS herstellen
# --------------------------------------------------------------
# API-Verbindung zu PVGIS       
    pvgis_url = "https://re.jrc.ec.europa.eu/api/v5_2/"     

    # Betrachtungszeitrahmen festlegen
        times15 = pd.date_range(
            "$(sunYear)-01-01 00:00:00",
            "$(sunYear)-12-31 23:45:00";
            freq="15min",
            tz=tz
        )

println("Lade PVGIS-ERA5 Daten für $(sunYear) …")

# Datenimport PVGIS Hourly (ERA5)
    kwargs = Dict{Symbol,Any}(
        :start => sunYear,
        :end => sunYear,
        :pvcalculation => false,
        :components => true,
        :usehorizon => true,
        :raddatabase => "PVGIS-ERA5",   
        :outputformat => "json",
        :map_variables => true,
        :url => pvgis_url,
        :surface_tilt => tilt,
        :surface_azimuth => azim,
    )

    data, meta = pvlib.iotools.get_pvgis_hourly(lat, lon; kwargs...)

data === nothing && error("PVGIS hat keine Daten geliefert (data === nothing)")

#
# 1) Simulation der PV-Stromproduktion
# --------------------------------------------------------------
# Zeitzone festlegen
    if pyconvert(Bool, data.index.tz == pybuiltins.None)
        data.index = data.index.tz_localize("UTC").tz_convert(tz)
    else
        data.index = data.index.tz_convert(tz)
    end

# 15-min Resampling der Daten
    data["temp_air"]   = data["temp_air"].ffill()
    data["wind_speed"] = data["wind_speed"].ffill()

    met15 = data.resample("15min").interpolate()
    met15 = met15.reindex(times15, method="nearest")

# POA global berechnen
    poa_global = (
        met15["poa_direct"] +
        met15["poa_sky_diffuse"] +
        met15["poa_ground_diffuse"]
    )

# Berücksichtigung der Zelltemperatur
    t_cell = pvlib.temperature.sapm_cell(
        poa_global,
        met15["temp_air"],
        met15["wind_speed"],
        a=-3.47, b=-0.0594, deltaT=10
    )

# PV-Watts DC zu AC
    poa_clipped = np.clip(poa_global, 0.0, np.inf)
    pdc = pvlib.pvsystem.pvwatts_dc(poa_clipped, t_cell, P_dc_W, gamma_pdc)
    pac = system_loss * n_inv * pdc   

# Nach Julia konvertieren
    time_strings = pyconvert(Vector{String}, times15.strftime("%Y-%m-%d %H:%M:%S"))
    time_jl = Dates.DateTime.(time_strings, dateformat"yyyy-mm-dd HH:MM:SS")

    pac_w = pyconvert(Vector{Float64}, pac.to_numpy())
    energy_kWh_15min = (pac_w ./ 1000.0) .* 0.25

    df = DataFrame(
        Timestamp = time_jl,
        PV_Power_W = pac_w,
        PV_Energy_kWh_15min = energy_kWh_15min
    )

println("Jahresertrag der PV-Anlage mit $(Int(P_dc_W/1000)) kWp: ", round(sum(df.PV_Energy_kWh_15min), digits=1), " kWh")

#
# 3) PV-Stromproduktion als Excel exportieren
# --------------------------------------------------------------
# Umwandeln jeder Julia‐Zeile in Dict
    records = [Dict(
        "Timestamp" => string(df.Timestamp[i]),
        "PV_Power_W" => df.PV_Power_W[i],
        "PV-Power [kWh]" => df.PV_Energy_kWh_15min[i]
    ) for i in 1:nrow(df)]

    df_py = pd.DataFrame(records) # Bauen eines Python-Pandas-DataFrame

# Reihenfolge der Spalten in Python anpassen (ohne Julia-Operatoren)
    cols = pylist(["Timestamp", "PV-Power [kWh]"])
    df_py = df_py.get(cols)

# Dateiablage
    base_path = @__DIR__
    file_name_PV = joinpath(base_path,"..", "data", "pv", "$(projekt)_PV-Power_$(Int(P_dc_W/1000))kWp_$(sunYear).xlsx" )

# Excel-Datei speichern
    df_py.to_excel(file_name_PV; index=false)

println("Excel gespeichert unter: ", file_name_PV)
