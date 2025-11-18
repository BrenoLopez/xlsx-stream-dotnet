namespace Report.Console.Contracts.Input;

public record InputRecord(
    string Description,
    byte Stars,
    string Title,
    string Category,
    DateTime CreatedAt,
    DateTime UpdatedAt,
    string Author,
    bool IsActive,
    int Priority,
    decimal Budget,
    Guid Identifier,
    string Tags,
    string Comments,
    double Rating,
    int Views
);
