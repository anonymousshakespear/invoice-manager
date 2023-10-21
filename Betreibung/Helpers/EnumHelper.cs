namespace Betreibung.Helpers; 
public static class EnumHelper {
    public static TEnum ParseEnum<TEnum>(string value) {
        return (TEnum)Enum.Parse(typeof(TEnum), value, true);
    }
}
