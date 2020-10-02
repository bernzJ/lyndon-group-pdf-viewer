import * as React from "react";

interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const Header = ({ title, logo, message }: HeaderProps) => (
  <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
    <img width="90" height="90" src={logo} alt={title} title={title} />
    <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
  </section>
);

export default Header;
