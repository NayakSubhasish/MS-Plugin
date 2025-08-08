import * as React from "react";
import PropTypes from "prop-types";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingBottom: "20px",
    paddingTop: "20px",
    backgroundColor: tokens.colorNeutralBackground3,
    width: "100%",
  },
  message: {
    fontSize: tokens.fontSizeHero1000,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForegroundStatic,
    marginTop: "10px",
  },
});

const Header = (props) => {
  const { title, message } = props;
  const styles = useStyles();
  const logo = "assets/logo-filled.webp";

  return (
    <section className={styles.welcome__header}>
      <Image 
        src={logo} 
        alt={title} 
        style={{
          maxWidth: '180px',
          width: '100%',
          height: 'auto',
          objectFit: 'contain'
        }}
      />
      <h1 className={styles.message}>{message}</h1>
    </section>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  message: PropTypes.string,
};

export default Header;
