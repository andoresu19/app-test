import { PersonCard } from "@microsoft/mgt-react";
import { useContext } from "react";
import { TeamsFxContext } from "../Context";

export function PersonCardGraphToolkit(props) {
  const { themeString } = useContext(TeamsFxContext);

  return (
    <div className="section-margin">
      <p>
        This example uses Graph Toolkit's&nbsp;
        <a
          href="https://docs.microsoft.com/en-us/graph/toolkit/components/person-card"
          target="_blank"
          rel="noreferrer"
        >
          person card component
        </a>{" "}
        with&nbsp;
        <a
          href="https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/providers/mgt-teamsfx-provider"
          target="_blank"
          rel="noreferrer"
        >
          TeamsFx provider
        </a>{" "}
        to show person card.
      </p>

      {!props.loading && props.error && (
        <div className="error">
          Failed to read your profile. Please try again later. <br /> Details:{" "}
          {props.error.toString()}
        </div>
      )}
      {!props.loading && !props.error && props.data && (
        <>
          <p>
            <strong>Personal info: </strong>
          </p>
          <div className={themeString === "default" ? "mgt-light" : "mgt-dark"}>
            <PersonCard personQuery="me" isExpanded={false}></PersonCard>
          </div>
          <div>
            <p>
              <strong>Organization: </strong>
              {props.data.organization.value[0].displayName}
            </p>
            <p>
              <strong>Users: </strong>
            </p>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(4, 1fr)",
                gap: "15px",
              }}
            >
              {props.data.users.value.map((item) => (
                <span
                  style={{
                    background: "#ffffff",
                    padding: "10px",
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                    textAlign: "center",
                    borderRadius: "5px",
                    boxShadow: "rgba(0, 0, 0, 0.15) 1.95px 1.95px 2.6px",
                    gap: "10px",
                  }}
                >
                  <strong>{item.displayName}</strong>
                  <span style={{ fontSize: "12px" }}>{item.id}</span>
                </span>
              ))}
            </div>
          </div>
        </>
      )}
    </div>
  );
}
