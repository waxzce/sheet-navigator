import { Input } from "@fluentui/react-components";
import { Search24Regular, Dismiss20Regular } from "@fluentui/react-icons";

type Props = {
  value: string;
  onChange: (next: string) => void;
};

export function SearchBar({ value, onChange }: Props): JSX.Element {
  return (
    <Input
      size="small"
      placeholder="Rechercher une feuille..."
      value={value}
      onChange={(_, data) => onChange(data.value)}
      contentBefore={<Search24Regular />}
      contentAfter={
        value ? (
          <Dismiss20Regular
            style={{ cursor: "pointer" }}
            onClick={() => onChange("")}
            aria-label="Effacer"
          />
        ) : undefined
      }
      style={{ width: "100%" }}
    />
  );
}
