from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application and PostgreSQL database configuration."""

    db_host: str = "10.50.100.28"
    db_port: int = 5432
    db_name: str = "db_snj_group"
    db_schema: str = "schm_snj"
    db_user: str = "postgres"
    db_password: str = ""
    app_port: int = 8000

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    @property
    def dsn(self) -> str:
        """Build PostgreSQL connection string."""
        return (
            f"host={self.db_host} port={self.db_port} dbname={self.db_name} "
            f"user={self.db_user} password={self.db_password}"
        )


settings = Settings()
